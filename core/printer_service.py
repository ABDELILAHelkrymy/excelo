import subprocess

def get_windows_printers():
    try:
        out = subprocess.check_output(
            ["powershell", "-Command", "Get-CimInstance Win32_Printer | Select-Object -ExpandProperty Name"],
            text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        return [p.strip() for p in out.splitlines() if p.strip()]
    except Exception:
        return []

def get_default_printer():
    try:
        out = subprocess.check_output(
            ["powershell", "-Command", "Get-CimInstance Win32_Printer | Where-Object Default -eq $true | Select-Object -ExpandProperty Name"],
            text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        return out.strip()
    except Exception:
        return None

def set_default_printer(printer_name):
    try:
        subprocess.run(["rundll32", "printui.dll,PrintUIEntry", "/y", "/n", printer_name], 
                       creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass
