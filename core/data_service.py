from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import pandas as pd


@dataclass
class DataOperationResult:
    dataframe: pd.DataFrame
    not_found: pd.DataFrame | None = None
    found: pd.DataFrame | None = None


def normalize_key(series: pd.Series) -> pd.Series:
    return series.astype("string").fillna("").str.strip()


class DataProcessor:
    def build_membership_set(self, df: pd.DataFrame, column: str) -> set[str]:
        self._ensure_column(df, column)
        return set(normalize_key(df[column]).tolist())

    def filter_by_membership(
        self,
        df: pd.DataFrame,
        column: str,
        membership: set[str],
        minus: bool,
    ) -> tuple[pd.DataFrame, set[str]]:
        self._ensure_column(df, column)
        normalized = normalize_key(df[column])
        mask = ~normalized.isin(membership) if minus else normalized.isin(membership)
        return df.loc[mask].copy(), set(normalized.tolist())

    def rows_not_in_membership(
        self,
        df: pd.DataFrame,
        column: str,
        membership: set[str],
    ) -> pd.DataFrame:
        self._ensure_column(df, column)
        normalized = normalize_key(df[column])
        return df.loc[~normalized.isin(membership)].copy()

    def subtract(
        self,
        df_a: pd.DataFrame,
        col_a: str,
        df_b: pd.DataFrame,
        col_b: str,
    ) -> pd.DataFrame:
        membership = self.build_membership_set(df_b, col_b)
        result, _ = self.filter_by_membership(df_a, col_a, membership, minus=True)
        return result

    def intersect(
        self,
        df_a: pd.DataFrame,
        col_a: str,
        df_b: pd.DataFrame,
        col_b: str,
        *,
        selected_a: list[str] | None = None,
        selected_b: list[str] | None = None,
        principal: str = "A",
        suffix: str = " (B)",
        include_not_found: bool = False,
    ) -> DataOperationResult:
        self._ensure_column(df_a, col_a)
        self._ensure_column(df_b, col_b)

        right_keys = normalize_key(df_b[col_b])
        left_keys = normalize_key(df_a[col_a])
        membership = set(right_keys.tolist())
        result = df_a.loc[left_keys.isin(membership)].copy()

        if selected_a is None:
            selected_a = []
        if selected_b is None:
            selected_b = []

        if selected_b:
            result = self._merge_selected_columns(
                left=result,
                left_key_column=col_a,
                right=df_b,
                right_key_column=col_b,
                selected_a=selected_a,
                selected_b=selected_b,
                principal=principal,
                suffix=suffix,
            )
        elif selected_a:
            keep = [column for column in selected_a if column in result.columns]
            if keep:
                result = result.loc[:, keep]

        not_found = None
        if include_not_found:
            not_found = df_b.loc[~right_keys.isin(set(left_keys.tolist()))].copy()
            if not len(not_found):
                not_found = None

        return DataOperationResult(dataframe=result, not_found=not_found)

    def extend(
        self,
        df_a: pd.DataFrame,
        df_b: pd.DataFrame,
        *,
        dedup: bool,
        dedup_key: str,
    ) -> pd.DataFrame:
        combined = pd.concat([df_a, df_b], ignore_index=True)
        if not dedup:
            return combined
        if dedup_key and dedup_key != "— aucune (toutes les colonnes) —" and dedup_key in combined.columns:
            return combined.drop_duplicates(subset=[dedup_key])
        return combined.drop_duplicates()

    def enrich(
        self,
        df_a: pd.DataFrame,
        col_a: str,
        df_b: pd.DataFrame,
        col_b: str,
        *,
        selected_b: list[str],
        suffix: str,
        join_mode: str,
        skip_empty: bool,
        fill_mode: str,
        fill_value: str,
    ) -> DataOperationResult:
        self._ensure_column(df_a, col_a)
        self._ensure_column(df_b, col_b)
        if not selected_b:
            raise ValueError("Sélectionnez au moins une colonne de B à ajouter.")

        left = df_a.copy()
        left["_match_key"] = normalize_key(left[col_a])

        right = df_b.loc[:, list(dict.fromkeys(selected_b + [col_b]))].copy()
        right["_match_key"] = normalize_key(df_b[col_b])

        if skip_empty:
            left = left.loc[left["_match_key"].ne("")].copy()
            right = right.loc[right["_match_key"].ne("")].copy()

        right = right.drop_duplicates(subset=["_match_key"])
        how = "left" if "LEFT" in join_mode else "inner"

        result = left.merge(right, on="_match_key", how=how, suffixes=("", suffix))
        result.drop(columns=["_match_key"], inplace=True, errors="ignore")

        found = left.merge(right, on="_match_key", how="inner", suffixes=("", suffix))
        found.drop(columns=["_match_key"], inplace=True, errors="ignore")
        if not len(found):
            found = None

        if fill_mode == "Chaîne vide":
            result = result.fillna("")
        elif fill_mode == "Valeur personnalisée":
            result = result.fillna(fill_value)

        return DataOperationResult(dataframe=result, found=found)

    def extract(
        self,
        df: pd.DataFrame,
        column: str,
        condition: str,
        value: str,
        *,
        case_sensitive: bool,
    ) -> pd.DataFrame:
        self._ensure_column(df, column)
        series = df[column].astype(str).str.strip()
        probe = value
        if not case_sensitive:
            series = series.str.lower()
            probe = value.lower()

        if condition == "égal à":
            mask = series == probe
        elif condition == "différent de":
            mask = series != probe
        elif condition == "contient":
            mask = series.str.contains(probe, regex=False, na=False)
        elif condition == "commence par":
            mask = series.str.startswith(probe, na=False)
        elif condition == "se termine par":
            mask = series.str.endswith(probe, na=False)
        elif condition == "supérieur à":
            try:
                mask = pd.to_numeric(df[column], errors="coerce") > float(value)
            except ValueError as exc:
                raise ValueError("'Supérieur à' nécessite une valeur numérique.") from exc
        elif condition == "inférieur à":
            try:
                mask = pd.to_numeric(df[column], errors="coerce") < float(value)
            except ValueError as exc:
                raise ValueError("'Inférieur à' nécessite une valeur numérique.") from exc
        elif condition == "est vide":
            mask = df[column].isna() | (df[column].astype(str).str.strip() == "")
        elif condition == "n'est pas vide":
            mask = ~(df[column].isna() | (df[column].astype(str).str.strip() == ""))
        else:
            raise ValueError(f"Condition inconnue : {condition}")

        return df.loc[mask].copy()

    def combine(
        self,
        frames: Iterable[pd.DataFrame],
        *,
        selected_columns: list[str],
        dedup: bool,
        dedup_key: str,
    ) -> pd.DataFrame:
        prepared_frames: list[pd.DataFrame] = []
        for frame in frames:
            if selected_columns:
                keep = [column for column in selected_columns if column in frame.columns]
                prepared_frames.append(frame.loc[:, keep])
            else:
                prepared_frames.append(frame.copy())

        if not prepared_frames:
            raise ValueError("Aucune donnée à combiner.")

        result = pd.concat(prepared_frames, ignore_index=True)
        if not dedup:
            return result
        if dedup_key and dedup_key != "— toutes les colonnes —" and dedup_key in result.columns:
            return result.drop_duplicates(subset=[dedup_key])
        return result.drop_duplicates()

    @staticmethod
    def _ensure_column(df: pd.DataFrame, column: str) -> None:
        if column not in df.columns:
            raise ValueError(f"Colonne '{column}' introuvable.")

    def _merge_selected_columns(
        self,
        *,
        left: pd.DataFrame,
        left_key_column: str,
        right: pd.DataFrame,
        right_key_column: str,
        selected_a: list[str],
        selected_b: list[str],
        principal: str,
        suffix: str,
    ) -> pd.DataFrame:
        right_columns = list(dict.fromkeys(selected_b + [right_key_column]))
        right_subset = right.loc[:, right_columns].copy()
        right_subset["_match_key"] = normalize_key(right[right_key_column])
        right_subset = right_subset.drop_duplicates(subset=["_match_key"])

        left = left.copy()
        left["_match_key"] = normalize_key(left[left_key_column])
        merge_suffixes = (suffix, "") if principal == "B" else ("", suffix)
        merged = left.merge(right_subset, on="_match_key", how="left", suffixes=merge_suffixes)
        merged.drop(columns=["_match_key"], inplace=True, errors="ignore")

        keep: list[str] = []
        if principal == "B":
            keep.extend(self._ordered_columns(selected_a, merged, suffix_first=True, suffix=suffix))
            keep.extend(self._ordered_columns(selected_b, merged, suffix_first=False, suffix=suffix))
        else:
            keep.extend(self._ordered_columns(selected_a, merged, suffix_first=False, suffix=suffix))
            keep.extend(self._ordered_columns(selected_b, merged, suffix_first=False, suffix=suffix))

        keep = list(dict.fromkeys(column for column in keep if column in merged.columns))
        return merged.loc[:, keep] if keep else merged

    @staticmethod
    def _ordered_columns(
        columns: list[str],
        df: pd.DataFrame,
        *,
        suffix_first: bool,
        suffix: str,
    ) -> list[str]:
        ordered: list[str] = []
        for column in columns:
            candidates = [f"{column}{suffix}", column] if suffix_first else [column, f"{column}{suffix}"]
            for candidate in candidates:
                if candidate in df.columns and candidate not in ordered:
                    ordered.append(candidate)
                    break
        return ordered