"""
Pure data operations for Excel Scripts Filter Tool.
All methods here run independently of the UI.
"""

import logging
import pathlib
import polars as pl


def cleanup_floats(df: pl.DataFrame) -> pl.DataFrame:
    """
    Utility: Casts Fastexcel Float64 columns inferred from integers back to Int64 globally.
    If a float column has no fractional parts, it is safely restored to Int64 to avoid saving as 5.0.
    """
    cast_exprs = []
    for c, dtype in zip(df.columns, df.dtypes):
        if dtype in (pl.Float32, pl.Float64):
            s = df[c].drop_nulls()
            if len(s) > 0 and s.eq(s.round(0)).all():
                cast_exprs.append(pl.col(c).cast(pl.Int64))
    if cast_exprs:
        return df.with_columns(cast_exprs)
    return df


def _coerce_fill_value(fill_value: str, dtype: pl.DataType):
    """Try to convert a custom fill value to the target Polars dtype."""
    try:
        return pl.Series("_fill", [fill_value], dtype=pl.Utf8).cast(dtype, strict=True).item()
    except Exception:
        return None


def op_minus_or_intersect(
    minus: bool,
    df_a: "pl.DataFrame | None", 
    df_b: "pl.DataFrame | None",
    mi_mode: str, 
    mi_a_col: str, 
    mi_b_col: str,
    mi_multi_files: dict, 
    mi_multi_key: str, 
    mi_nf: bool,
    mi_a_col_vars: dict, 
    mi_b_col_vars: dict,
    mi_suffix: str, 
    mi_principal: str
) -> tuple[pl.DataFrame, "pl.DataFrame | None"]:
    """
    Perform subtraction (minus=True) or intersection (minus=False) between A and B.
    Returns: (result_df, notfound_df)
    """
    notfound = None

    # ── Resolve A side (single or multi) ────────────────────────────
    if mi_mode == "Plusieurs A × 1 B":
        if not mi_multi_files:
            raise ValueError("Ajoutez au moins un Fichier A.")
        if df_b is None:
            raise ValueError("Le Fichier B n'est pas chargé.")
        if mi_b_col not in df_b.columns:
            raise ValueError(f"Colonne '{mi_b_col}' introuvable dans le Fichier B.")
        multi_key = mi_multi_key

        b_keys = df_b.select(pl.col(mi_b_col).cast(pl.Utf8).str.strip_chars().alias("_key"))

        frames: list = []
        join_how = "anti" if minus else "semi"
        for path, current_df in mi_multi_files.items():
            if multi_key not in current_df.columns:
                raise ValueError(
                    f"Colonne '{multi_key}' introuvable dans "
                    f"{pathlib.Path(path).name}.")
            df_with_key = current_df.with_columns(pl.col(multi_key).cast(pl.Utf8).str.strip_chars().alias("_key"))
            res = df_with_key.join(b_keys, on="_key", how=join_how).drop("_key")
            frames.append(res)

        result = pl.concat(frames, how="vertical_relaxed") if frames else pl.DataFrame()

        # Not-found: B keys absent from ALL A files
        if not minus and mi_nf:
            all_a_keys = pl.concat([
                df.select(pl.col(multi_key).cast(pl.Utf8).str.strip_chars().alias("_key"))
                for df in mi_multi_files.values()
            ], how="vertical_relaxed")
            db_b_full = df_b.with_columns(pl.col(mi_b_col).cast(pl.Utf8).str.strip_chars().alias("_key"))
            nf = db_b_full.join(all_a_keys, on="_key", how="anti").drop("_key")
            if not nf.is_empty():
                notfound = nf
        return result, notfound

    # ── Resolve B side (single or multi) ────────────────────────────
    if mi_mode == "1 A × Plusieurs B":
        if df_a is None:
            raise ValueError("Le Fichier A n'est pas chargé.")
        if not mi_multi_files:
            raise ValueError("Ajoutez au moins un Fichier B.")
        if mi_a_col not in df_a.columns:
            raise ValueError(f"Colonne '{mi_a_col}' introuvable dans le Fichier A.")
        multi_key = mi_multi_key

        ldf_a = df_a.with_columns(pl.col(mi_a_col).cast(pl.Utf8).str.strip_chars().alias("_key"))

        # Union all B keys
        all_b_keys_list = []
        for path, current_df in mi_multi_files.items():
            if multi_key not in current_df.columns:
                raise ValueError(
                    f"Colonne '{multi_key}' introuvable dans "
                    f"{pathlib.Path(path).name}.")
            all_b_keys_list.append(current_df.select(pl.col(multi_key).cast(pl.Utf8).str.strip_chars().alias("_key")))

        b_keys = pl.concat(all_b_keys_list, how="vertical_relaxed").unique()

        join_how = "anti" if minus else "semi"
        result = ldf_a.join(b_keys, on="_key", how=join_how).drop("_key")

        # Not-found: union of B keys absent from A
        if not minus and mi_nf:
            a_keys = ldf_a.select("_key").unique()
            nf_frames: list = []
            for bdf in mi_multi_files.values():
                bdf_k = bdf.with_columns(pl.col(multi_key).cast(pl.Utf8).str.strip_chars().alias("_key"))
                nf = bdf_k.join(a_keys, on="_key", how="anti").drop("_key")
                if not nf.is_empty():
                    nf_frames.append(nf)
            if nf_frames:
                nf = pl.concat(nf_frames, how="vertical_relaxed")
                if not nf.is_empty():
                    notfound = nf
        return result, notfound

    # ── Standard (1 A × 1 B) ────────────────────────────────────────
    if df_a is None:
        raise ValueError("Le Fichier A n'est pas chargé.")
    if df_b is None:
        raise ValueError("Le Fichier B n'est pas chargé.")
    if mi_a_col not in df_a.columns:
        raise ValueError(f"Colonne '{mi_a_col}' introuvable dans le Fichier A.")
    if mi_b_col not in df_b.columns:
        raise ValueError(f"Colonne '{mi_b_col}' introuvable dans le Fichier B.")

    ldf_a = df_a.with_columns(pl.col(mi_a_col).cast(pl.Utf8).str.strip_chars().alias("_key"))
    b_keys = df_b.select(pl.col(mi_b_col).cast(pl.Utf8).str.strip_chars().alias("_key"))

    join_how = "anti" if minus else "semi"
    result = ldf_a.join(b_keys, on="_key", how=join_how).drop("_key")

    # ── Intersect: column selection & merge ─────────────────────────────
    if not minus and (mi_a_col_vars or mi_b_col_vars):
        sel_a = [c for c, v in mi_a_col_vars.items() if v]
        sel_b = [c for c, v in mi_b_col_vars.items() if v]

        sfx = mi_suffix or " (B)"
        principal = mi_principal

        if sel_b:
            # Merge to bring selected B columns into result
            db_b_sel = df_b.with_columns(pl.col(mi_b_col).cast(pl.Utf8).str.strip_chars().alias("_mk"))
            b_keep = []
            for c in sel_b + ["_mk"]:
                if c not in b_keep:
                    b_keep.append(c)
            db_b_sel = db_b_sel.select(b_keep).unique(subset=["_mk"], maintain_order=True)

            result = result.with_columns(pl.col(mi_a_col).cast(pl.Utf8).str.strip_chars().alias("_mk"))

            if principal == "B":
                # Rename A-side clashing columns to carry the suffix, keep B as primary
                clash = [c for c in sel_a if c in sel_b]
                if clash:
                    result = result.rename({c: f"{c}{sfx}" for c in clash if c in result.columns})
                result = result.join(db_b_sel, on="_mk", how="left")
            else:
                result = result.join(db_b_sel, on="_mk", how="left", suffix=sfx)

            result = result.drop("_mk", strict=False)

        # Keep only selected columns
        if sel_a or sel_b:
            keep: list = []
            for c in sel_a:
                candidate = f"{c}{sfx}" if principal == "B" else c
                fallback  = c if principal == "B" else f"{c}{sfx}"
                if candidate in result.columns and candidate not in keep:
                    keep.append(candidate)
                elif fallback in result.columns and fallback not in keep:
                    keep.append(fallback)
            for c in sel_b:
                candidate = c if principal == "B" else c
                fallback  = f"{c}{sfx}"
                if candidate in result.columns and candidate not in keep:
                    keep.append(candidate)
                elif fallback in result.columns and fallback not in keep:
                    keep.append(fallback)
            if keep:
                result = result.select(keep)

    # ── Intersect: compute "not found" rows from B ──────────────────────
    if not minus and mi_nf:
        db_b = df_b.with_columns(pl.col(mi_b_col).cast(pl.Utf8).str.strip_chars().alias("_key"))
        a_keys = ldf_a.select("_key").unique()
        nf = db_b.join(a_keys, on="_key", how="anti").drop("_key")
        if not nf.is_empty():
            notfound = nf

    return result, notfound


def op_extend(
    df_a: "pl.DataFrame | None", 
    df_b: "pl.DataFrame | None",
    ext_dedup: bool, 
    ext_key: str,
) -> pl.DataFrame:
    """Fusionner (A U B)"""
    if df_a is None:
        raise ValueError("Le Fichier A n'est pas chargé.")
    if df_b is None:
        raise ValueError("Le Fichier B n'est pas chargé.")
    combined = pl.concat([df_a, df_b], how="diagonal_relaxed")
    if ext_dedup:
        if ext_key and ext_key != "— aucune (toutes les colonnes) —" and ext_key in combined.columns:
            combined = combined.unique(subset=[ext_key], maintain_order=True)
        else:
            combined = combined.unique(maintain_order=True)
    return combined


def op_enrich(
    df_a: "pl.DataFrame | None", 
    df_b: "pl.DataFrame | None",
    en_a_col: str, 
    en_b_col: str,
    en_cols: dict, 
    en_mode: str,
    en_suffix: str, 
    en_skip_empty: bool,
    en_fill: str, 
    en_fill_val: str,
) -> tuple[pl.DataFrame, "pl.DataFrame | None"]:
    """
    Enrichir (A ← B): LEFT or INNER JOIN on key columns, bringing selected B columns.
    Returns (result_df, en_found_df).
    """
    if df_a is None:
        raise ValueError("Le Fichier A n'est pas chargé.")
    if df_b is None:
        raise ValueError("Le Fichier B n'est pas chargé.")

    if en_a_col not in df_a.columns:
        raise ValueError(f"Colonne '{en_a_col}' introuvable dans le Fichier A.")
    if en_b_col not in df_b.columns:
        raise ValueError(f"Colonne '{en_b_col}' introuvable dans le Fichier B.")

    # Columns from B the user wants to add
    sel_b = [c for c, v in en_cols.items() if v]
    if not sel_b:
        raise ValueError("Sélectionnez au moins une colonne de B à ajouter.")

    sfx = en_suffix or " (B)"

    # Prepare A
    ldf_a = df_a.with_columns(pl.col(en_a_col).cast(pl.Utf8).str.strip_chars().alias("_mk"))

    # Prepare B subset — keep only selected columns + merge key
    ldf_b = df_b.with_columns(pl.col(en_b_col).cast(pl.Utf8).str.strip_chars().alias("_mk"))

    # Optionally skip rows with empty / NaN keys
    if en_skip_empty:
        ldf_a = ldf_a.filter((pl.col("_mk").is_not_null()) & (pl.col("_mk") != "") & (pl.col("_mk") != "nan"))
        ldf_b = ldf_b.filter((pl.col("_mk").is_not_null()) & (pl.col("_mk") != "") & (pl.col("_mk") != "nan"))

    b_keep = []
    for c in sel_b + ["_mk"]:
        if c not in b_keep:
            b_keep.append(c)
    df_b_sel = ldf_b.select(b_keep).unique(subset=["_mk"], maintain_order=True)

    # Determine join type
    how = "left" if "LEFT" in en_mode else "inner"

    result = ldf_a.join(df_b_sel, on="_mk", how=how, suffix=sfx) # type: ignore
    result = result.drop("_mk", strict=False)

    # Fill empty cells produced by unmatched rows on the imported B columns.
    fill_exprs = []
    for col in sel_b:
        target_col = f"{col}{sfx}" if col in df_a.columns else col
        if target_col not in result.columns:
            continue

        if en_fill == "Chaîne vide":
            fill_exprs.append(
                pl.when(pl.col(target_col).is_null())
                .then(pl.lit(""))
                .otherwise(pl.col(target_col).cast(pl.Utf8))
                .alias(target_col)
            )
            continue

        if en_fill == "Valeur personnalisée":
            dtype = result.schema[target_col]
            coerced_value = _coerce_fill_value(en_fill_val, dtype)
            if coerced_value is not None:
                fill_exprs.append(pl.col(target_col).fill_null(coerced_value).alias(target_col))
            else:
                fill_exprs.append(
                    pl.when(pl.col(target_col).is_null())
                    .then(pl.lit(en_fill_val))
                    .otherwise(pl.col(target_col).cast(pl.Utf8))
                    .alias(target_col)
                )

    if fill_exprs:
        result = result.with_columns(fill_exprs)
    # "Laisser vide" → keep NaN as-is

    # Compute "found in B": A rows that matched B
    en_found = ldf_a.join(df_b_sel, on="_mk", how="inner", suffix=sfx)
    en_found = en_found.drop("_mk", strict=False)
    en_found_result = en_found if not en_found.is_empty() else None

    return result, en_found_result


def op_extract(df: "pl.DataFrame | None", ex_filters: list) -> pl.DataFrame:
    """Extraire (Filtrer A)"""
    if df is None:
        raise ValueError("Le Fichier A n'est pas chargé.")
    if not ex_filters:
        return df

    final_mask = None

    for i, fil in enumerate(ex_filters):
        col  = fil["col"]
        if col not in df.columns or col.startswith("—"):
            if len(ex_filters) == 1 and col.startswith("—"):
                raise ValueError("Veuillez charger un fichier avant d'extraire.")
            raise ValueError(f"Colonne '{col}' introuvable dans le Fichier A.")

        cond = fil["cond"]
        val  = fil["val"]
        case = fil["case"]

        if cond == "égal à":
            if case:
                mask = pl.col(col).cast(pl.Utf8) == val
            else:
                mask = pl.col(col).cast(pl.Utf8).str.to_lowercase() == val.lower()
        elif cond == "différent de":
            if case:
                mask = pl.col(col).cast(pl.Utf8) != val
            else:
                mask = pl.col(col).cast(pl.Utf8).str.to_lowercase() != val.lower()
        elif cond == "contient":
            if case:
                mask = pl.col(col).cast(pl.Utf8).str.contains(val, literal=True)
            else:
                mask = pl.col(col).cast(pl.Utf8).str.to_lowercase().str.contains(val.lower(), literal=True)
        elif cond == "commence par":
            if case:
                mask = pl.col(col).cast(pl.Utf8).str.starts_with(val)
            else:
                mask = pl.col(col).cast(pl.Utf8).str.to_lowercase().str.starts_with(val.lower())
        elif cond == "se termine par":
            if case:
                mask = pl.col(col).cast(pl.Utf8).str.ends_with(val)
            else:
                mask = pl.col(col).cast(pl.Utf8).str.to_lowercase().str.ends_with(val.lower())
        elif cond == "supérieur à":
            try:
                numeric_val = float(val)
            except ValueError:
                raise ValueError("'Supérieur à' nécessite une valeur numérique.")
            try:
                mask = pl.col(col).cast(pl.Float64, strict=False) > numeric_val
            except Exception:
                raise ValueError(f"La colonne '{col}' ne peut pas être comparée numériquement.")
        elif cond == "inférieur à":
            try:
                numeric_val = float(val)
            except ValueError:
                raise ValueError("'Inférieur à' nécessite une valeur numérique.")
            try:
                mask = pl.col(col).cast(pl.Float64, strict=False) < numeric_val
            except Exception:
                raise ValueError(f"La colonne '{col}' ne peut pas être comparée numériquement.")
        elif cond == "est vide":
            mask = pl.col(col).is_null() | (pl.col(col).cast(pl.Utf8).str.strip_chars() == "")
        elif cond == "n'est pas vide":
            mask = pl.col(col).is_not_null() & (pl.col(col).cast(pl.Utf8).str.strip_chars() != "")
        else:
            raise ValueError(f"Condition inconnue : {cond}")

        if i == 0:
            final_mask = mask
        else:
            logic = fil["logic"]
            if logic == "ET":
                final_mask = final_mask & mask
            elif logic == "OU":
                final_mask = final_mask | mask

    if final_mask is None:
        return df

    return df.filter(final_mask)


def op_combine(
    cb_files: dict, 
    cb_dedup: bool,
    cb_dedup_key: str, 
    cb_cols: dict,
) -> pl.DataFrame:
    """Combiner: stack multiple files vertically, keeping selected columns."""
    if not cb_files:
        raise ValueError("Ajoutez au moins un fichier à combiner.")

    sel_cols = [c for c, v in cb_cols.items() if v]

    frames: list = []
    for path, sheets in cb_files.items():
        for df in sheets.values():
            if sel_cols:
                keep = [c for c in sel_cols if c in df.columns]
                frames.append(df.select(keep))
            else:
                frames.append(df.clone())

    if not frames:
        raise ValueError("Aucune donnée à combiner.")

    result = pl.concat(frames, how="diagonal_relaxed")

    if cb_dedup:
        if cb_dedup_key and cb_dedup_key != "— toutes les colonnes —" and cb_dedup_key in result.columns:
            result = result.unique(subset=[cb_dedup_key], maintain_order=True)
        else:
            result = result.unique(maintain_order=True)

    return result
