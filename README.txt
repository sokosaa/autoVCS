To use this program, make sure autoVCS.exe and default_VCS_template.pptx are in the same folder/directory. Those are the only 2 files required.

How to use / Function Descriptions:
1. enter working path (where you want to work in your file directory) (reccomend clean empty folder or VCS only folder)
2. enter the part numbers you want to make a VCS for in a vertical list (one number per line)
3. click create folders (a folder for each part number will be created in your working directory/folder
4. put the part images into their respective folders manually
5. upload custom PPTX template (optional) (or/also, you can edit the default template if you want)
6. open, editc and close each PPTX file manually
7. click PPTX to PDF (all pptx files you just made will be saved as pdfs as well)
8. upload to DDS manually
9. enter vendor code (all parts you are doing at once must be from the same vendor to use this next button)
10. click move to focus (MAKE SURE ALL PARTS YOU ARE DOING AT ONCE ARE FROM THE SAME VENDOR) All pdf files will be moved to their appropriate location in //focus without overwriting anything there (anythere there will be renamed to previousName_old.pdf)
11. a list of paths will show so you can confirm if you wish where the pdf files went


To update this program with changes, the required files (in the same directory) are:
README.txt
vcstools.py
gui.py
icon.ico
default_VCS_template.pptx
create_exe_command.txt

How to update program:
1. Edit vcstools.py and gui.py to your satisfaction for new / edited features, etc. (Make sure to use a version control system like git if you know how or simply just keep a copy of the two pyhton files incase changes break the program and you need to revert to a previous state.) 
2. After changes have been made to the program, it can be test run by just running the gui.py file. If all is well, a new exe file can be generated. 
3. First, to keep clean workspace, ensure you have a directory with and only with the 6 required files previously listed. (If you are using the same folder you can freely delete the 'autoVCS.spec' file and the 'build' folder. Also move the previous autoVCS.exe elsewhere or delete it, but make sure to hold on to default_VCS_template.pptx)
4. Then, open a terminal (Command Prompt works) and navigate to the directory containing the program files (example command: "cd C:\Users\example\autoVCSeg")
5. Finally, paste the command found in 'create_exe_command.txt' and press enter. 
6. After a short while, you will have an exe file 'autoVCS.exe' in a 'dist' folder.
7. To use this program, move the 'default_VCS_template.pptx' file into the 'dist' folder. You can move this folder anywhere you want including sending it to another pc.


