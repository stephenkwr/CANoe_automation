from pywinauto import Desktop
import os, time, winreg, shutil
from pywinauto.timings import *
from pywinauto import Desktop, timings

# find the UTAS execution engine path on the machine. Default install path is D: drive. However, not all computers have :D drives so
# this code attempts to find the executable using the uninstall path as uTAS registers in windows registry on download
def _try(path): 
    return path if path and os.path.exists(path) else None

def find_UTAS_Execution_Engine_Path() -> str | None:
    uninstall_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\uTAS5",       winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\uTAS5", winreg.KEY_WOW64_32KEY),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\uTAS5",       winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\uTAS5", winreg.KEY_WOW64_32KEY),
    ]

    install_root = None
    exe_from_uninstall = None

    for hive, subkey, view in uninstall_keys:
        try:
            with winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ | view) as k:
                # Prefer InstallLocation
                try:
                    install_root = winreg.QueryValueEx(k, "InstallLocation")[0]
                except FileNotFoundError:
                    install_root = None
                # UninstallString may point to "...sys\\uninstall.exe" or a helper
                try:
                    uninstall_str = winreg.QueryValueEx(k, "UninstallString")[0]
                except FileNotFoundError:
                    uninstall_str = None
                if uninstall_str:
                    exe_path = uninstall_str.strip('"')
                    if exe_path.lower().endswith(".exe"):
                        # If it's inside ...\sys\ , go up to root
                        sys_dir = os.path.dirname(exe_path)
                        guessed_root = os.path.dirname(sys_dir)
                        if os.path.isdir(guessed_root):
                            install_root = install_root or guessed_root
        except FileNotFoundError:
            continue

    # Try canonical bin path from install_root
    if install_root:
        p = os.path.join(install_root, "bin", "ExecutionEngine.exe")
        hit = _try(p)
        if hit: 
            return hit

    # PATH lookup
    hit = shutil.which("ExecutionEngine.exe")
    if hit:
        return os.path.abspath(hit)

    # Common roots
    for base in (r"D:\uTAS5\bin", r"C:\uTAS5\bin", r"C:\Program Files\uTAS5\bin", r"C:\Program Files (x86)\uTAS5\bin"):
        hit = _try(os.path.join(base, "ExecutionEngine.exe"))
        if hit:
            return hit

    # Program Files scan (shallow)
    for pf in (r"C:\Program Files", r"C:\Program Files (x86)", r"D:\Program Files", r"D:\Program Files (x86)"):
        candidate = os.path.join(pf, "uTAS5", "bin", "ExecutionEngine.exe")
        hit = _try(candidate)
        if hit:
            return hit

    return None

# Check for the OTC window login window if it exist and block until it closes either by cancelling or actually logging in
def wait_for_OTC_Login(appear_timeout=30, close_timeout=300, fail_if_not_closed=True):
    """
    Block if the 'OTC Login Interface' dialog appears.
    Returns:
        True  -> dialog appeared and was closed
        False -> dialog never appeared within appear_timeout
    Raises:
        RuntimeError -> dialog appeared but did not close within close_timeout and fail_if_not_closed=True
    """
    desktop = Desktop(backend="uia")

    # Match title loosely (case-insensitive, extra spaces ok)
    selector = {
        "title_re": r"(?i)\s*OTC\s+Login\s+Interface\s*",
        "control_type": "Window"
    }

    # 1) Wait for the dialog to APPEAR
    try:
        timings.wait_until(
            appear_timeout, 0.2,
            lambda: desktop.window(**selector).exists()
        )
    except timings.TimeoutError:
        # Never showed up - nothing to block on
        return False

    # It exists now
    dlg = desktop.window(**selector)
    try:
        dlg.wait("exists enabled visible ready", timeout=5)
        dlg.set_focus()
        # dlg.activate()  # optional
    except Exception:
        pass

    # 2) Block until it DISAPPEARS
    try:
        dlg.wait_not("exists", timeout=close_timeout)
        return True
    except timings.TimeoutError:
        if fail_if_not_closed:
            raise RuntimeError(
                f"OTC Login dialog did not close within {close_timeout} seconds."
            )
        return False

# take in as input of .txt file and read it line by line and convert the numrical value to that of a percentage upon 255
# eg value is 128. 128 / 255 * 100 is 50%. For sound level percentage for Suzuki. Writes the values to a out.txt
def to_Percentage_Of_255_From_Txt(data_file):
    with open(data_file, "r") as infile, open("out.txt", "w") as outfile:
        for line in infile:
            s = line.strip()                # 1. remove “\n”
            if not s:                       
                continue                    # skip empty lines
            val = int(s)                    # 2. decimal string → int
            pct = (val / 255) * 100         # 3. convert to percentage
            out = f"{pct:.2f}"              # 4. limit to 2 decimal place
            outfile.write(out + "\n")       # 5. write to outfile

# take in a value, check if its an numeric number in either int, float or string. if yes, it convert it to percentage out 255
# eg value is 128. 128 / 255 * 100 is 50%. For sound level percentage for Suzuki. Writes the values to a out.txt
def to_Percentage_Of_255(value, as_str=True):
    if value is None:
        return "0.00" if as_str else 0.0
    try:
        val = float(str(value).strip())   # handles '123', '123.0', 123, 123.0
    except (TypeError, ValueError):
        return "0.00" if as_str else 0.0

    pct = (val / 255.0) * 100.0
    return f"{pct:.2f}" if as_str else round(pct, 2)

# function to convert a int to hex val to string 
# eg input int = 10 
# converted hex = 0x0A
# final return = 0A
def convert_to_hex_string_without_prefix(volume):
    hexed_val = hex(volume)
    # print(f"{volume=}, {hexed_val=}, {hexed_val[2:]}")
    return hexed_val[2:]