import os
from tkinter import Tk, Label, Button, filedialog
from tkinter.ttk import Combobox

# Currently not being used. 

# Use only if user is using this. If HW team is using this, not needed
def GUI_For_User():
    # storage for results
    results = {
        "output_excel": None,
        "config_excel": None,
        "cfg_file": None,
        "duration": None,
        "repeats": None
    }

    def pick_output_excel():
        path = filedialog.askopenfilename(
            title="Select Output Excel file",
            filetypes=[("Excel files","*.xlsx;*.xlsm;*.xls")],
            defaultextension=".xlsx"
        )
        if path:
            results["output_excel"] = path
            btn_output.config(text=os.path.basename(path))

    def pick_config_excel():
        path = filedialog.askopenfilename(
            title="Select Config Excel file",
            filetypes=[("Excel files","*.xlsx;*.xlsm;*.xls")],
            defaultextension=".xlsx"
        )
        if path:
            results["config_excel"] = path
            btn_config.config(text=os.path.basename(path))

    def pick_cfg():
        path = filedialog.askopenfilename(
            title="Select Simulation CFG file",
            filetypes=[("Config files","*.cfg")],
            defaultextension=".cfg"
        )
        if path:
            results["cfg_file"] = path
            btn_cfg.config(text=os.path.basename(path))

    def on_enter():
        # read UI controls into results
        results["duration"] = int(cb_duration.get())
        results["repeats"]  = int(cb_repeats.get())
        root.quit()

    root = Tk()
    root.title("Suzuki menu")
    root.geometry("500x350")
    root.resizable(False, False)

    # Output Excel picker
    Label(root, text="Output Excel:").grid(row=0, column=0, padx=10, pady=8, sticky="e")
    btn_output = Button(root, text="Choose .xlsx (optional)", command=pick_output_excel, width=25)
    btn_output.grid(row=0, column=1, pady=8)

    # Config Excel picker
    Label(root, text="Config Excel:").grid(row=1, column=0, padx=10, pady=8, sticky="e")
    btn_config = Button(root, text="Choose .xlsx", command=pick_config_excel, width=25)
    btn_config.grid(row=1, column=1, pady=8)

    # Simulation .cfg picker
    Label(root, text="Simulation file (.cfg):").grid(row=2, column=0, padx=10, pady=8, sticky="e")
    btn_cfg = Button(root, text="Choose .cfg", command=pick_cfg, width=25)
    btn_cfg.grid(row=2, column=1, pady=8)

    # Duration dropdown (1 or 10)
    Label(root, text="Duration (s):").grid(row=3, column=0, padx=10, pady=8, sticky="e")
    cb_duration = Combobox(root, values=[1, 10], width=23, state="readonly")
    cb_duration.current(0)
    cb_duration.grid(row=3, column=1, pady=8)

    # Repeats dropdown (1–10)
    Label(root, text="Repeats:").grid(row=4, column=0, padx=10, pady=8, sticky="e")
    cb_repeats = Combobox(root, values=list(range(1,11)), width=23, state="readonly")
    cb_repeats.current(0)
    cb_repeats.grid(row=4, column=1, pady=8)

    # Enter button
    Button(root, text="Enter", command=on_enter, width=20, bg="lightgray")\
        .grid(row=5, column=0, columnspan=2, pady=15)

    root.mainloop()
    root.destroy()

    return (
        results["output_excel"],
        results["config_excel"],
        results["cfg_file"],
        results["duration"],
        results["repeats"]
    )

# only ask for config file. May be needed(?) 
def GUI_For_HW_Team():
    # storage for results
    result = {"config_excel": None}

    def pick_config_excel():
        path = filedialog.askopenfilename(
            title="Select Config Excel file",
            filetypes=[("Excel files","*.xlsx;*.xlsm;*.xls")],
            defaultextension=".xlsx"
        )
        if path:
            result["config_excel"] = path
            btn_config.config(text=os.path.basename(path))

    def on_enter():
        # read UI controls into results
        root.quit()

    root = Tk()
    root.title("Suzuki menu")
    root.geometry("300x350")
    root.resizable(False, False)

    # Config Excel picker
    Label(root, text="Config Excel:").grid(row=1, column=0, padx=10, pady=8, sticky="e")
    btn_config = Button(root, text="Choose .xlsx", command=pick_config_excel, width=25)
    btn_config.grid(row=1, column=1, pady=8)

    # Enter button
    Button(root, text="Enter", command=on_enter, width=20, bg="lightgray")\
        .grid(row=5, column=0, columnspan=2, pady=15)

    root.mainloop()
    root.destroy()

    return result["config_excel"]