# haVBScript
# haVBScripts

### GetShortcutTarget.vbs  
This VBscript copies the target file from a shortcut into the  
directory where the shortcut resides.   
Optionally, the copy of the target may be renamed.  
Usage:      
Copy the script named *GetShortcutTarget.vbs* to  
`C:\Users\...\AppData\Roaming\Microsoft\Windows\SendTo`.  
In Windows Explorer click the right mouse-button on the  
shortcut(s) or on the folder containing your shortcut(s),  
and, in the context menu choose *SendTo*   
and select *GetShortcutTarget.vbs*.  
This will run the script and process the shortcuts appropriately.  

![screenshot1](document/image/Steamhammer01.jpg)    

![screenshot1a](document/image/Steamhammer02.jpg)    

### ChangeShortcut.vbs  
This VBscript adjusts the target path and working directory in shortcuts.     
Usage:      
Copy the script named *ChangeShortcut.vbs* to  
`C:\Users\...\AppData\Roaming\Microsoft\Windows\SendTo`.  
In Windows Explorer click the right mouse-button on the  
folder containing your shortcut(s), and, in the context menu  
choose *SendTo* and select *ChangeShortcut.vbs*.  
This will run the script and process the shortcuts appropriately.  

### BuildVersion.vbs / SetVersion.vbs 
Build a string showing date and time formatted as a version string.  
Return the version string as a function in a generated C++ source module.  
Compile the module `haCryptBuildTime.cpp` holding the version information  
and link `haCryptBuildTime.obj` with the c++ project.    
