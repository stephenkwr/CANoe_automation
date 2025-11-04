from common_modules.RPA            import *
from common_modules.GUI            import *
from common_modules.UTAS_wrapper   import *
from common_modules.File_IO        import *
from common_modules.HelperFunc     import *
from pywinauto.application import ProcessNotFoundError
from datetime import datetime, timedelta

# for suzuki

# strings
output_path_name = "Output.xlsx" # output file name
simulation_file_path = "" # path to be provided by user
cfg_file_path = "" # to be provided by user in excel form
telegram_msg_play = "31 01 fe 23 "
telegram_msg_stop = "31 02 fe 23"
telegram_initialise = "10 60"

# integers
duration = 12 # 1 by default, options in GUI will be 10 or 1 seconds. Duration the sound will be played for
repeats = 1 # 3 by default, can range from 1 - 10. Number of repeats per sound
security_Key = 0 # value is 0, 1, 2 depending on user requirement. Security key to choose the security access

# for mitsubishi specific case
start_vol = 10 # beginning volume to test
end_vol = 255 # end volume to test
# 3 as of 24/9/2025 for tolerance
tolerance = 3 # how much variance from recorded highest db. this is dependant on specification given by user. If user requires multiple tolerances, take the smallest tolerances (eg. 1 sound has tolerance of 3 while another is 5. Tolerance for all sounds is 3)

# helper function to initialise or reinitialise. Known to have issues when running where sound stops running as error and need to 10 60 again
def check_last_received_response(UTAS):
    strResp = UTAS.send_command("get_env", ["Diag_LastResp_Text", "str"])
    while strResp != "OK":
        UTAS.send_command("set_env", ["Diag_FreeDiagTelegram_Data", telegram_initialise]) # send index of sound to be played
        UTAS.send_command("toggle_env", ["Diag_FreeDiagTelegram_Btn", "200"]) # start sound playing
        time.sleep(2)
        strResp = UTAS.send_command("get_env", ["Diag_LastResp_Text", "str"])


if __name__ == "__main__":
    start_time = datetime.now()
    start = time.perf_counter()
    print(f"Start:   {start_time:%Y-%m-%d %H:%M:%S}")
    UTAS_Execution_Engine_Path = find_UTAS_Execution_Engine_Path() # find the path on the machine for UTAS execution engine
    assert UTAS_Execution_Engine_Path is not None 
    app = Application()
    try:
        app.connect(path=UTAS_Execution_Engine_Path) # try to find execution engine if already running
    except(ProcessNotFoundError, RuntimeError):
        app = app.start(UTAS_Execution_Engine_Path, wait_for_idle=False) # if not running, launch

    # if given to hardware team, uncomment this
    sounds_To_Play, output_path_name, simulation_file_path, duration, repeats, security_Key = read_Config_File_For_HW_Team() # read from config file for parameters
    print(f"{output_path_name=}, {cfg_file_path=}, {simulation_file_path=}, {duration=}, {repeats=}, {security_Key=}") # print the variables 

    UTAS = UtasWrapper()
    RPA_automation = RPA()

    output_wb = open_Output_Excel(path=output_path_name, test_name=simulation_file_path) # excel file result is to be written into. Creates it if it does not exist

    UTAS.load_project_settings(UTAS_PROJECT_PATH) # load the initial proj. May change based on requirement
    UTAS.send_command("save_setting", ['"CANoe.cfg_set.cfg_group.SimulationConfigPath.value"', simulation_file_path]) # set the simulation file path in the prj setting dynamically.
    # UTAS.send_command("get_setting", ['"CANoe.cfg_set.cfg_group.SimulationConfigPath.value"']) # check if the file path has been set

    # change this for mitsubishi based on cfg file
    UTAS.send_command("open_simulation") # open the simulation file (.cfg)
    UTAS.send_command("delay", ["1000"]) # wait for the simulation file to open 
    UTAS.send_command("start_simulation") # start the simulation

    UTAS.send_command("set_env", ["Env_TesterPresent", "1"]) # click on tester present to prevent exiting diagnostic mode. If not on, default behaviour is to exit diagnostic mode after 5 seconds of no input
    UTAS.send_command("set_env", ["Diag_FreeDiagTelegram_Data", telegram_initialise]) # send index of sound to be played
    UTAS.send_command("toggle_env", ["Diag_FreeDiagTelegram_Btn", "200"]) # start sound playing
    check_last_received_response(UTAS=UTAS)
    
    no_Sounds = len(sounds_To_Play) # number of sounds to be played
    for row, (index, level) in enumerate(sounds_To_Play, start = 1): # iterate through all the sounds. Rows indicate how many rows will be in the excel starting from row 1 (excel is 1 based indexing)
        diag_msg_idx_no_vol = telegram_msg_play + convert_to_hex_string_without_prefix(index) # initial message with out volume
        col = 1
        for vol in range(1, end_vol - start_vol, 10): # begin at index 1 as open excel is 1 based index (eg first cell is 1, 1). Repeats 245 times. plays the sound from sound level 10 to 255
            current_vol = vol - 1 + start_vol # calculate what is the current volume to play at for this iteration
            current_diag_msg = diag_msg_idx_no_vol + " " + convert_to_hex_string_without_prefix(current_vol) # this is the final telegram diagnostic message to be passed to the simulation
            print(f"********************{row}/{no_Sounds} sounds played. Playing sound index {index} at sound level {current_vol}. ********************")
            check_last_received_response(UTAS=UTAS)
            UTAS.send_command("set_env", ["Diag_FreeDiagTelegram_Data", current_diag_msg]) # input the message to play the sound
            UTAS.send_command("toggle_env", ["Diag_FreeDiagTelegram_Btn", "200"]) # enter the message to play the sound
            RPA_automation.measure_Sound(Rec_duration=duration) # start measurement, let the duration elapse before stopping
            UTAS.send_command("set_env", ["Diag_FreeDiagTelegram_Data", telegram_msg_stop]) # input the message to stop the sound
            UTAS.send_command("toggle_env", ["Diag_FreeDiagTelegram_Btn", "200"]) # enter the message to stop the sound
            RPA_automation.save_CSV(iter = vol, Rec_duration = duration) # save the recorded CSV for data extraction
            highest_recorded_dB = RPA_automation.process_CSV(iter = vol, Rec_duration = duration) # get the highest measured dB for sound played
            write_Into_Cell(wb = output_wb, row = row, col = col, data = highest_recorded_dB) # write into output excel file
            if (level - tolerance) <= highest_recorded_dB <= (level + tolerance): # check if the recorded sound is within tolerance of given volume
                bold_text(wb = output_wb, row = row, col = col) # bold the text in the excel for easy identification
            col += 1
            output_wb.save(output_path_name)
    
    end = time.perf_counter()
    end_time = datetime.now()
    elapsed = end - start
    print(f"Test completed! {no_Sounds}/{no_Sounds} sounds played.")
    print(f"Start:   {start_time:%Y-%m-%d %H:%M:%S}")
    print(f"End:     {end_time:%Y-%m-%d %H:%M:%S}.")
    print(f"Elapsed: {elapsed:.3f} s  ({timedelta(seconds=elapsed)})")
    input("Press ENTER to exit…")