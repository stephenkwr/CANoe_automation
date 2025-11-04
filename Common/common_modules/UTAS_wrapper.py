import clr, sys, os, winreg

def find_utas_install_root():
    """
    Finds the “InstallLocation” for the uTAS5 uninstall key by scanning
    all Uninstall subkeys for DisplayName containing “uTAS”, just as we
    did for the lib folder—but this time we return the parent folder.
    """
    uninstall = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    hives = [
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),
        (winreg.HKEY_CURRENT_USER,  winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_CURRENT_USER,  winreg.KEY_WOW64_32KEY),
    ]
    for hive, flag in hives:
        try:
            with winreg.OpenKey(hive, uninstall, 0, winreg.KEY_READ | flag) as ukey:
                for i in range(winreg.QueryInfoKey(ukey)[0]):
                    sub = winreg.EnumKey(ukey, i)
                    with winreg.OpenKey(ukey, sub) as sk:
                        try:
                            name = winreg.QueryValueEx(sk, "DisplayName")[0]
                        except FileNotFoundError:
                            continue
                        if "uTAS" not in name:
                            continue
                        install_loc = winreg.QueryValueEx(sk, "InstallLocation")[0]
                        if os.path.isdir(install_loc):
                            return install_loc
        except FileNotFoundError:
            continue

    # fallback if we still haven’t found it
    for drive in ("C:", "D:", "E:", "F:"):
        candidate = rf"{drive}\uTAS5"
        if os.path.isdir(candidate):
            return candidate

    raise RuntimeError("Could not locate uTAS5 installation root")

# at module top, after you’ve already found utas_lib:
utas_root = find_utas_install_root()

UTAS_PROJECT_PATH = os.path.join(
    utas_root,
    "Projects",
    "Test",          # or whatever subfolder you actually want
    "v01.00.00"
)

def find_utas_lib_folder():
    uninstall = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    hives = [
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),
        (winreg.HKEY_CURRENT_USER,  winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_CURRENT_USER,  winreg.KEY_WOW64_32KEY),
    ]
    for hive, flag in hives:
        try:
            with winreg.OpenKey(hive, uninstall, 0, winreg.KEY_READ | flag) as ukey:
                n = winreg.QueryInfoKey(ukey)[0]
                for i in range(n):
                    sub = winreg.EnumKey(ukey, i)
                    with winreg.OpenKey(ukey, sub) as sk:
                        try:
                            name = winreg.QueryValueEx(sk, "DisplayName")[0]
                        except FileNotFoundError:
                            continue
                        if "uTAS" not in name:
                            continue
                        # now we know it’s the right key:
                        loc = winreg.QueryValueEx(sk, "InstallLocation")[0]
                        lib = os.path.join(loc, "lib")
                        if os.path.isdir(lib):
                            print("Found UTAS lib at", lib)
                            return lib
        except FileNotFoundError:
            continue

    # fallback if registry fails:
    for drive in ["C:", "D:", "E:"]:
        lib = fr"{drive}\uTAS5\lib"
        if os.path.isdir(lib):
            # print("Falling back to", lib)
            return lib

    return None

utas_lib = find_utas_lib_folder()
if utas_lib:
    # make sure pythonnet can load the assemblies
    if hasattr(os, "add_dll_directory"):
        os.add_dll_directory(utas_lib)   # Python 3.8+
    sys.path.append(utas_lib)            # for any pure-Python bits


clr.AddReference("uTAS.API")
clr.AddReference("uTAS.Communication.ExecEngineComAPI")
import uTAS.Communication.ExecEngineComAPI as EECOM

class UtasWrapper():
    def __init__(self, clientName = "PythonClient", port = 8888):
        try:
            self.EECOM_OBJ = EECOM.ExecEngineCommunicationClient(port, clientName)
            # connect to the ExecutionEngine context
            self.EECOM_OBJ.Connect().ConfigureAwait(False).GetAwaiter().GetResult()
        except Exception as e:
            self.error_log("Connecting to uTAS Error: {}".format(e))

    def load_project_settings(self, project_path):
        response = self.EECOM_OBJ.SendCmdRequest("load_project_settings", [project_path, "default", "default"])
        if response.get_Err().get_Description() is not None:
            self.error_log("Loading project Settings error: {}".format(response.get_Err().get_Description()))
            assert response.get_Err().get_Description() == None, "Loading project Settings error"
            return False

    def send_command(self, command: str, param: list = []):
        result = None
        try:
            response = self.EECOM_OBJ.SendCmdRequest(command, param)

            errorDesc = response.get_Err().get_Description()
            errorFlag = False if errorDesc is None else True
            
            if errorFlag:
                raise Exception(errorDesc)
            else:
                result = response.get_Result()
            print(f"{command=}, {param=}, {result=}")
        except Exception as e:
            self.error_log("send_command Error when sending {}: {}".format(command, e))
        finally:
            return result
    def error_log(self, msg):
        print(msg)