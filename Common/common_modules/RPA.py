from pywinauto import Application, Desktop
import os, time, sys, winreg
import pandas as pd
import numpy as np
import tempfile, atexit, shutil

# variables for ARTA
ARTA_config_file = r"audioconfig"
arta_exe_loc = r"C:\Program Files (x86)\ArtaSoftware\Arta.exe"
CSV_Dir = "Logged_CSV_Val"

# helper function for printMenuItems to print individual menu elements
def __printFormatedString(inString, rightSpacing):
    string_length = len(inString) + rightSpacing
    string_revised = inString.rjust(string_length)
    print(string_revised)

# Recursive function to print all menu elements from dialog menus
def printMenuItems(menu_item, mainWindow, menuName, spacing):
    # Open the Menu
    menu_item.invoke()
    time.sleep(0.5)
    #Access the menubar via application
    menu = mainWindow.child_window(title="Application", auto_id="MenuBar", control_type="MenuBar")
    # Access current Titled "title" menu item
    curMenu = menu.child_window(title=menuName, control_type="Menu")

    # iterate over the menu Items
    for menuItem in curMenu.children():
        # Get Menu Name from legacy properties
        menuName = menuItem.legacy_properties().get(u'Name')
        # Print formatted String based upon spacing
        __printFormatedString(inString=menuName, rightSpacing=spacing*2)
        # If this is a submenu and recusively call this function
        if menuItem.legacy_properties().get(u'ChildId') == 0:
            # recursive call
            printMenuItems(menu_item=menuItem, mainWindow=mainWindow, menuName=menuName, spacing=spacing*2)
            # press ESC to remove the menu
            mainWindow.type_keys("{ESC}")

class RPA():
    # constructor, also initialises arta and sets it up for sound measurement
    def __init__(self):
        arta_exe_loc = self.__find_arta_via_registry()
        assert os.path.exists(arta_exe_loc)  # check if executable exist, if not exit.

        # ---- READ-ONLY RESOURCE BASE (for bundled assets like .cal) ----
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # running as EXE: bundled resources live under _MEIPASS
            self.base = os.path.abspath(sys._MEIPASS)
        else:
            # running as script: Common/ is parent of modules/
            self.base = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

        # ---- RUNTIME TEMP OUTPUTS (auto-clean) ----
        # Place per-run temp next to the EXE when frozen, else in current working dir
        if getattr(sys, 'frozen', False):
            _temp_root = os.path.dirname(sys.executable)
        else:
            _temp_root = os.getcwd()

        # Create a private per-run temp directory (e.g., <exe_dir>\AA_xxx\Logged_CSV_Val)
        self._tmpdir = tempfile.mkdtemp(prefix="AA_", dir=_temp_root)
        self.write_base = os.path.join(self._tmpdir, CSV_Dir)
        os.makedirs(self.write_base, exist_ok=True)

        # Optional: expose for callers/tests
        self.csv_dir = self.write_base

        # Auto-cleanup at process exit (unless debugging)
        def _cleanup_tmp():
            keep = os.environ.get("AA_KEEP_TEMP") == "1"
            if keep:
                return
            try:
                shutil.rmtree(self._tmpdir, ignore_errors=True)
            except Exception:
                pass

        atexit.register(_cleanup_tmp)
        # ---- end runtime temp outputs ----

        # Locate the ARTA config (.cal) from bundled resources (read-only)
        cal_file = os.path.join(self.base, "AudioConfigArta", ARTA_config_file + ".cal")
        assert os.path.exists(cal_file)  # check if the file exist, else exit

        app = Application(backend="uia").start(arta_exe_loc)  # launch ARTA via UIA
        win = app.window(title_re="(?i).*arta.*")             # match ARTA main window

        # 1) Open Setup
        win.type_keys("%S")  # Alt+S
        time.sleep(0.1)
        win.type_keys("A")   # Audio devices

        # 2) Audio Devices Setup window
        audio_subwin = win.child_window(title="Audio Devices Setup", control_type="Window")
        audio_subwin.child_window(title="Load setup", control_type="Button").invoke()

        # 3) File-Open dialog: enter full path to .cal and open
        audio_subwin.wait("visible")
        open_dlg = Desktop(backend="uia").window(title_re="Open", control_type="Window", top_level_only=False)
        open_dlg.wait("visible", timeout=5)

        fn_combo = open_dlg.child_window(auto_id="1148", control_type="ComboBox")
        fn_edit  = fn_combo.child_window(control_type="Edit")
        fn_edit.set_edit_text(cal_file)
        open_dlg.type_keys("%O")  # Alt+O (Open)

        # Confirm and return to main
        audio_subwin.child_window(title="OK", control_type="Button").invoke()

        # 4) Tools → Integrating (SPL meter)
        win.type_keys("%T")  # Alt+T
        time.sleep(0.1)
        win.type_keys("I")
        time.sleep(0.1)

        # Focus SPL meter window
        IPL_subwin = win.child_window(title="SPL meter (noname.spl)", control_type="Window")

        # Integration speed: Fast
        IPL_subwin.child_window(auto_id="1261", control_type="ComboBox").wrapper_object().select("Fast")

        # Range → Set → SPL graph setup
        IPL_subwin.child_window(auto_id="1215", control_type="Button").invoke()
        IPL_subwin.wait("visible")
        SPL_graph_setup = IPL_subwin.child_window(title="SPL graph setup", control_type="Window")

        # Untoggle LEQ, LSlow, LImpulse, LPeak if needed
        leq_chk = SPL_graph_setup.child_window(auto_id="1442", control_type="CheckBox").wrapper_object()
        if leq_chk.get_toggle_state() == 1:
            leq_chk.iface_toggle.Toggle()
        lslow_chk = SPL_graph_setup.child_window(auto_id="1444", control_type="CheckBox").wrapper_object()
        if lslow_chk.get_toggle_state() == 1:
            lslow_chk.iface_toggle.Toggle()
        limpulse_chk = SPL_graph_setup.child_window(auto_id="1445", control_type="CheckBox").wrapper_object()
        if limpulse_chk.get_toggle_state() == 1:
            limpulse_chk.iface_toggle.Toggle()
        lpeak_chk = SPL_graph_setup.child_window(auto_id="1446", control_type="CheckBox").wrapper_object()
        if lpeak_chk.get_toggle_state() == 1:
            lpeak_chk.iface_toggle.Toggle()

        # Confirm SPL graph setup
        SPL_graph_setup.child_window(auto_id="1", control_type="Button").invoke()
        self.IPL_subwin = IPL_subwin

    # find the registry 
    def __find_arta_via_registry(self):
        uninstall_subpath = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        # Try both hives × both registry views
        combos = [
            (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),
            (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),
            (winreg.HKEY_CURRENT_USER,  winreg.KEY_WOW64_64KEY),
            (winreg.HKEY_CURRENT_USER,  winreg.KEY_WOW64_32KEY),
        ]
        for hive, view_flag in combos:
            try:
                with winreg.OpenKey(
                    hive,
                    uninstall_subpath,
                    0,
                    winreg.KEY_READ | view_flag
                ) as uninstall_key:
                    count = winreg.QueryInfoKey(uninstall_key)[0]
                    for i in range(count):
                        subkey_name = winreg.EnumKey(uninstall_key, i)
                        with winreg.OpenKey(uninstall_key, subkey_name) as sub:
                            # 1) look for DisplayName containing “Arta”
                            try:
                                name = winreg.QueryValueEx(sub, "DisplayName")[0]
                            except FileNotFoundError:
                                continue
                            if "Arta" not in name:
                                continue

                            # 2) prefer InstallLocation\Arta.exe
                            try:
                                install_loc = winreg.QueryValueEx(sub, "InstallLocation")[0]
                                path = os.path.join(install_loc, "Arta.exe")
                                if os.path.exists(path):
                                    return path
                            except FileNotFoundError:
                                pass

                            # 3) fallback: parse DisplayIcon (may include “,0” suffix)
                            try:
                                icon = winreg.QueryValueEx(sub, "DisplayIcon")[0]
                                path = icon.split(",")[0]
                                if os.path.exists(path):
                                    return path
                            except FileNotFoundError:
                                pass
            except FileNotFoundError:
                # hive/view not present — skip it
                continue

        # Not found in registry → maybe fallback to Program Files standard path
        for base in (r"C:\Program Files", r"C:\Program Files (x86)"):
            candidate = os.path.join(base, "ArtaSoftware", "Arta.exe")
            if os.path.exists(candidate):
                return candidate

        return None

    # to extract peak measured dB in saved CSV file recorded
    def process_CSV(self, iter, Rec_duration):
        # read from runtime temp folder (auto-cleaned at exit)
        CSV_loc = os.path.join(self.write_base, f"spl-{Rec_duration}s-log-{iter}.csv")
        CSV_path = os.path.normpath(CSV_loc)

        # 1) Read absolutely everything as data (no headers)
        df = pd.read_csv(CSV_path, header=None, skip_blank_lines=True, on_bad_lines='skip', dtype=str, keep_default_na=False)

        # 2) Build a mask of “which cells contain the substring LAFmax?”
        contains_lafmax = df.map(lambda s: 'LAFmax' in s)

        # 3) Find the first True in that mask
        coords = np.argwhere(contains_lafmax.values)
        if coords.size == 0:
            raise ValueError("Couldn’t find any cell containing LAFmax")

        row, col = coords[0]                # first occurrence
        raw = df.iat[row, col+1]            # the cell immediately to its right

        # 4) Strip off “ dB” (or anything after the number) and convert to float
        lafmax_db = float(raw.strip().split()[0])

        print(f"{lafmax_db=}")
        return lafmax_db

    # To save CSV file from ARTA. Rec_duration is 0.1, 1, or 10 seconds accordingly
    def save_CSV(self, iter, Rec_duration):
        # write into runtime temp folder (auto-cleaned at exit)
        CSV_loc = os.path.join(self.write_base, f"spl-{Rec_duration}s-log-{iter}.csv")
        CSV_path = os.path.normpath(CSV_loc)

        self.IPL_subwin.type_keys("%F")  # Alt+F (File)
        time.sleep(0.1)
        self.IPL_subwin.type_keys("E")   # Export
        if Rec_duration < 1:
            self.IPL_subwin.type_keys("C")       # CSV for 100ms recording
        elif Rec_duration >= 1 and Rec_duration < 10:
            self.IPL_subwin.type_keys("{S 2}")   # CSV for 1s recording
            time.sleep(0.1)
            self.IPL_subwin.type_keys("{ENTER}")
        else:
            self.IPL_subwin.type_keys("V")       # CSV for 10s recording

        save_dlg = Desktop(backend="uia").window(title_re="Save As", control_type="Window", top_level_only=False)
        time.sleep(10)  # allow dialog to appear (tune if needed)

        fn_combo = save_dlg.child_window(auto_id="FileNameControlHost", control_type="ComboBox")
        fn_edit  = fn_combo.child_window(auto_id="1001", control_type="Edit")

        # If same file exists from earlier run, delete it first
        if os.path.exists(CSV_path):
            os.remove(CSV_path)

        # Key in the file path to save as
        fn_edit.set_edit_text(CSV_path)

        # Save
        save_dlg.type_keys("%S")  # Alt+S (Save)
        time.sleep(2)  # tune if UI is slow

    # measure the sound for a duration then stop recording once the time has elapsed
    def measure_Sound(self, Rec_duration):
        self.IPL_subwin.child_window(title="Record/Reset", control_type="Button").wrapper_object().invoke()
        if 0 < Rec_duration <= 1:
            time.sleep(1)
        else:
            time.sleep(12)
        self.IPL_subwin.child_window(title="Stop", control_type="Button").wrapper_object().invoke()
