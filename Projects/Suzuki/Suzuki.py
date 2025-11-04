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
soundtune_string = "SoundTune_SoundNo" # to select index for tones
soundtune_vol = "SoundTune_SoundVolume_new" # to set volume of tones
soundtune_play = "SoundTune_PlaySound" # to play the tone
soundtune_stop = "SoundTune_StopSound" # to stop the tone

soundvoice_string = "SoundTune_VoiceNo" # to select index for voices
soundvoice_vol = "SoundTune_VoiceVolume_new" # to set volume for voices
soundvoice_play = "SoundTune_PlayVoice" # to play the tone
soundvoice_stop = "SoundTune_StopVoice" # to stop the tone

current_index_box = soundtune_string 
current_vol_box = soundtune_vol
current_play_butt = soundtune_play
current_stop_butt = soundtune_stop

# integers
duration = 1 # 1 by default, options in GUI will be 10 or 1 seconds. Duration the sound will be played for
repeats = 3 # 3 by default, can range from 1 - 10. Number of repeats per sound
security_Key = 0 # value is 0, 1, 2 depending on user requirement. Security key to choose the security access

security_Access_Diag_Success = frozenset(["I/O Control By Local Identifier: Positive Response", # Response from successfully playing sound in SoundTune panel 
                                          "Init Diagnostic Session: Positive Response", # Response from clicking VDO producion
                                          "Security Access: Positive Response", # Response from successfully connecting after clicking VDO programming 
                                          "Access Granted ! !"
                                        ])

security_Access_Diags_Fail = frozenset(["Security Access: RC ($24) not yet defined", # Response unsuccessfully connecting after clicking VDO programming
                                        "Unknown Service:Security Access Denied", # Response from unsuccessfully playing sound in SoundTune panel
                                         "Security Access: Invalid Key" # Response unsuccessfully connecting after clicking VDO programming
                                        ])

if __name__ == "__main__":
    start_time = datetime.now()
    start = time.perf_counter()
    print(f"Start:   {start_time:%Y-%m-%d %H:%M:%S}")
    UTAS_Execution_Engine_Path = find_UTAS_Execution_Engine_Path() # find the path on the machine for UTAS execution engine
    print("UTAS_Execution_Engine_Path =", UTAS_Execution_Engine_Path)
    assert UTAS_Execution_Engine_Path is not None 
    app = Application()
    try:
        app.connect(path=UTAS_Execution_Engine_Path) # try to find execution engine if already running
    except(ProcessNotFoundError, RuntimeError):
        app = app.start(UTAS_Execution_Engine_Path, wait_for_idle=False) # if not running, launch

    # # if given to user, uncomment this
    # output_path_name, cfg_file_path, simulation_file_path, duration, repeats =  GUI_For_User() # singular gui function to get input from user
    # if not output_path_name:
    #     output_path_name = "Output.xlsx"
    # sounds_To_Play = read_Config_File_With_User_Input(config_file=cfg_file_path) # read information from config file. it returns a list of tuple pairs.

    # if given to hardware team, uncomment this
    sounds_To_Play, output_path_name, simulation_file_path, duration, repeats, security_Key = read_Config_File_For_HW_Team() # read from config file for parameters
    print(f"{output_path_name=}, {cfg_file_path=}, {simulation_file_path=}, {duration=}, {repeats=}, {security_Key=}") # print the variables 

    UTAS = UtasWrapper()
    RPA_automation = RPA()

    output_wb = open_Output_Excel(path=output_path_name, test_name=simulation_file_path) # excel file result is to be written into. Creates it if it does not exist

    UTAS.load_project_settings(UTAS_PROJECT_PATH) # load the initial proj. May change based on requirement
    UTAS.send_command("save_setting", ['"CANoe.cfg_set.cfg_group.SimulationConfigPath.value"', simulation_file_path]) # set the simulation file path in the prj setting dynamically.
    # UTAS.send_command("get_setting", ['"CANoe.cfg_set.cfg_group.SimulationConfigPath.value"']) # check if the file path has been set

    UTAS.send_command("open_simulation") # open the simulation file (.cfg)
    UTAS.send_command("delay", ["1000"]) # wait for the simulation file to open 
    UTAS.send_command("start_simulation") # start the simulation

    UTAS.send_command("set_env", ["MAIN_eTerminal15", "1"]) # to turn on the CAN tx on/off in main panel. 

    UTAS.send_command("toggle_env", ["TESTER_eDMInitVdoProduction", "200"]) # click on VDO production
    UTAS.send_command("set_env", ["TESTER_eProdType", str(security_Key)]) # set the security access key based on user requirement
    time.sleep(2)
    UTAS.send_command("toggle_env", ["TESTER_eDMSecAccessVdo", "200"]) # click on VDO programming

    ################ for if cyber security and OTC login is required. If not required, comment out from the below line till the end of OTC code chunk comment #############################
    if wait_for_OTC_Login(): # if OTC appears and require user input to login, it will block until OTC is achieved
        security_Access_Display = UTAS.send_command("get_env", ["TESTER_eDMSecAccess_Display", "str"])
        security_Access_Diag_Status_Line = UTAS.send_command("get_env", ["TESTER_eDMDiagStatusLine", "str"])
        print(f"{security_Access_Display=}, {security_Access_Diag_Status_Line=}")

        start = 0
        total_Wait_Time = 5

        while security_Access_Display not in security_Access_Diag_Success and security_Access_Diag_Status_Line not in security_Access_Diag_Success:
            time.sleep(3)
            start += 3
            if start >= total_Wait_Time:
                UTAS.send_command("toggle_env", ["TESTER_eDMSecAccessVdo", "200"]) # click on VDO programming again as it tends to timeout and require clicking it again to grant access
                break
        time.sleep(10)
    ############################# end of OTC code chunk here ############################################
    no_Sounds = len(sounds_To_Play)
    for row, (index, level) in enumerate(sounds_To_Play, start = 1):
        if index < 0:
            current_index_box = soundvoice_string
            current_vol_box = soundvoice_vol
            current_play_butt = soundvoice_play
            current_stop_butt = soundvoice_stop
            continue
        # write_Into_Cell(wb = output_wb, row = row, col = 1, data = row)
        for col in range(1, repeats+1):
            percent_Level = to_Percentage_Of_255(value=level, as_str=True) # convert value given in config from upon 255 to percentage
            print(f"********************{row}/{no_Sounds} sounds played. Playing sound index {index} at sound level {percent_Level}. Repeated: {col}/{repeats} ********************")
            UTAS.send_command("set_env", [current_index_box, str(index)]) # send index of sound to be played
            UTAS.send_command("set_env", [current_vol_box, str(percent_Level)]) # send sound level of sound to be played
            UTAS.send_command("toggle_env", [current_play_butt, "200"]) # start sound playing
            RPA_automation.measure_Sound(Rec_duration=duration) # start measurement, let the duration elapse before stopping
            UTAS.send_command("toggle_env", [current_stop_butt, "200"]) # stop sound playing
            RPA_automation.save_CSV(iter = col, Rec_duration = duration) # save the recorded CSV
            highest_recorded_dB = RPA_automation.process_CSV(iter = col, Rec_duration = duration) # get the highest measured dB for sound played
            write_Into_Cell(wb = output_wb, row = row, col = col, data = highest_recorded_dB) # write into output excel file
            output_wb.save(output_path_name)
        calculate_And_Write_Average(output_wb, row = row, number_of_repeats = repeats) # calculate average for the output row and append to the end
        output_wb.save(output_path_name)
    end = time.perf_counter()
    end_time = datetime.now()
    elapsed = end - start
    print(f"Test completed! {no_Sounds}/{no_Sounds} sounds played.")
    print(f"Start:   {start_time:%Y-%m-%d %H:%M:%S}")
    print(f"End:     {end_time:%Y-%m-%d %H:%M:%S}.")
    print(f"Elapsed: {elapsed:.3f} s  ({timedelta(seconds=elapsed)})")
    input("Press ENTER to exit…")