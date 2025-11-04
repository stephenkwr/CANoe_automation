LINK TO DOCUMENTS: \\igdb001\didc0005\41_TechInfo\ME\0b_Architecture_Design\20_MECHATRONIC_ACOUSTIC\AcousticLab\ASSAT (link may have expired. Contact Leonardus - UIDC0339 for access to files)

To build:  
pyinstaller "projects/*project folder*/*project file*.py" `
  --onefile `
  --name "*project name*" `
  --paths "Common" `
  --add-data "${PWD}\Common\AudioConfigArta;AudioConfigArta" `
  --hidden-import "tkinter" `
  --distpath "Projects/*project folder*/dist" `
  --workpath "Projects/*project folder*/build" `
  --specpath "Projects/*project folder*"


EG. For suzuki project, the relevant UTAS flow is defined in "Suzuki.py" and we wish to name the executable "Suzuki_Auto.exe". The console command is then
pyinstaller "projects/Suzuki/Suzuki.py" `
  --onefile `
  --name "Suzuki_auto" `
  --paths "Common" `
  --add-data "${PWD}\Common\AudioConfigArta;AudioConfigArta" `
  --hidden-import "tkinter" `
  --distpath "Projects/Suzuki/dist" `
  --workpath "Projects/Suzuki/build" `
  --specpath "Projects/Suzuki"

The AudioConfigArta folder containing the .cal file that is used to initialise ARTA. You can generate a new .cal file by opening ARTA, setting the specifications you want and saving the configurations into
a .cal file. Save it in the "AudioConfigArta" folder and name it "audioconfig.cal" within the common folder. 

Generated executable will be located in a generated "dist" folder in each project folder.