REM --------------------------------------------------------------------------- 
REM              --------------------------------
REM             | VBScript: "ChangeShortcut.vbs" |
REM              --------------------------------
REM
REM  $_SCOPE  : Windows XP SP 3 (Language German)
REM                
REM  $_GENERAL:This VBscript adjusts the target path and working directory
REM            in shortcuts. 
REM             Usage:
REM              Copy the script named "ChangeShortcut.vbs" to
REM              "C:\Dokumente und Einstellungen\UserName\SendTo".
REM              In Windows Explorer click the right mouse-button on the
REM              folder containing your shortcut(s), and, in the context menu
REM              choose "Senden an" and select "ChangeShortcut.vbs".
REM              This will run the script and process the shortcuts appropriately.
REM              The initial idea was provided by Helmut Rohrbeck www.helmrohr.de.
REM
REM  $_NOTES  : IMPORTANT NOTE (from Microsoft):
REM              The Create Shortcut command truncates the source 
REM              path folder names to eight characters.
REM
REM     When you create shortcuts and specify a long file name in the target path,
REM      the path is truncated if the hard disk for the target does not exist.
REM       For example, create a shortcut with the following target:
REM       "J:\Mydirectory\Myapplication.exe"
REM       If drive "J" does not exist, the path is truncated to:
REM       "J:\Mydirect\Mypplica.exe" 
REM     
REM     This problem can occur because the shell cannot determine whether
REM      the hard disk supports long file names, so the path is truncated
REM      to be acceptable to all file systems. 
REM     
REM     Microsoft has confirmed that this is a problem in the Microsoft
REM      products that are listed at the beginning of this article. 
REM     
REM     This problem may be observed when you use any of the following
REM      methods to create shortcuts: 
REM      1) The Systems Management Server (SMS) Installer Create Shortcut Method
REM      2) The VBScript Create Shortcut Method
REM      3) The IShellLink Interface Method
REM     
REM     --------------------------------------
REM     2) The VBScript Create Shortcut method
REM     A sample VBScript that demonstrates the problem:
REM 
REM       set WshShell = WScript.CreateObject("WScript.Shell")
REM       set oShellLink = WshShell.CreateShortcut("d:\" & "\Long filename Shortcut .lnk")
REM       oShellLink.TargetPath = "j:\my long directory\myapplication.exe"
REM       oShellLink.WindowStyle = 1
REM       oShellLink.Hotkey = "CTRL+SHIFT+F"
REM       oShellLink.Description = "Long Filename Shortcut"
REM       oShellLink.Save
REM                     
REM     When you run this script and drive "J" does not exist,
REM      you can observe the created shortcut, but the target path is:
REM      "J:\My_long_\Myapplic.exe"
REM     
REM     NOTE: Any characters that are not normally supported by file systems
REM      that do not want long file names, such as the space character,
REM      are replaced by the underscore symbol "_".
REM     
REM     To work around this problem, you can use the subst command
REM      to point drive "J" to a local hard disk: 
REM     
REM       set WshShell = WScript.CreateObject("WScript.Shell")
REM       Dim ret
REM       'subst a drive to make the mapping work
REM       ret = WshShell.Run ("cmd /c subst j: c:\", 0, TRUE)
REM       set oShellLink = WshShell.CreateShortcut("d:\" & "\Long filename Shortcut .lnk")
REM       oShellLink.TargetPath = "j:\my long directory\myapplication.exe"
REM       oShellLink.WindowStyle = 1
REM       oShellLink.Hotkey = "CTRL+SHIFT+F"
REM       oShellLink.Description = "Long Filename Shortcut"
REM       oShellLink.Save
REM       'remove the subst
REM       ret = WshShell.Run ("cmd /c subst j: /d", 0, TRUE)
REM                     
REM     This command points drive J to drive C.
REM      If drive C supports long file names,
REM       the command creates a shortcut with the following target path:
REM       "J:\My long directory\Myapplication.exe"
REM
REM  $_DATE   : 30.07.2013
REM  $_AUTHOR : HelmutAltmann
REM --------------------------------------------------------------------------- 

Dim Tp, Wd, ret, TargetDriveLetter
Const TIMEOUT_1s = 1 

' Normally only the letter of the drive, hosting the target files, must be adjusted. 
' The drive's letter most likely change and is not (or no longer) consistent
' with a shortcut, when changing the plug-order of external (USB) drives.
' Enter a new drive letter for repacement in the current shortcut(s)
'
TargetDriveLetter =  UCase(InputBox ("Neuen Laufwerk-Buchstaben für Verknüpfung eingeben:", _
                                     "Verknüpfungen in einem Verzeichnis ändern.", _
                                     "C"))    
If TargetDriveLetter = "" Then
    WScript.Quit         ' Inputbox was Cancelled - Abort to Operation System.                                     
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

On Error Resume Next  ' Turn error handling on (may also hide some syntax errors) 

' NOTE: When you create shortcuts and specify a long file name
'       in the target path, the path folder name is truncated to eight chars, 
'       if the hard disk for the target does not exist.
'
'       Microsoft has confirmed that this is a problem.
' 
' WORKAROUND: Subst the (possibly) non-existent drive to prevent truncation of 
'             path folder names to eight chars, when running this script.
'
ret = WshShell.Run ("cmd /c subst " & TargetDriveLetter & ": C:\", 0, TRUE) 

Set oFSO    = CreateObject("Scripting.FileSystemObject")
Set cmdArgs = WScript.Arguments
Set oFolder = oFSO.GetFolder(cmdArgs(0))  ' Get 1st commandline argument, that is 
                                  '  the folder's path passed to 'sendto' 

For Each oFile In oFolder.Files
    If UCase(Right(oFile.Name, 4)) = ".LNK" Then      ' Only shortcuts will be processed

        Set oSC = WshShell.CreateShortcut(oFile.Path) ' Connect the shortcut

        ' Change the target's path in the shortcut:
        '  Replace(string, "X:\Old\Path", "X:\New\Path")
        ' Note: We currently most likely want to change the drive letter(s).
        '       However the replacements below could be tailored to meet any
        '       special requirements. 
        '
        Wd = oSC.WorkingDirectory
        Wd = Replace(Wd, "C:\", TargetDriveLetter & ":\")  
        Wd = Replace(Wd, "F:\", TargetDriveLetter & ":\")  
        Wd = Replace(Wd, "G:\", TargetDriveLetter & ":\")  
        Wd = Replace(Wd, "M:\", TargetDriveLetter & ":\")  
        Wd = Replace(Wd, "N:\", TargetDriveLetter & ":\")  

        Tp = oSC.TargetPath
        Tp = Replace(Tp, "C:\", TargetDriveLetter & ":\")  
        Tp = Replace(Tp, "F:\", TargetDriveLetter & ":\")  
        Tp = Replace(Tp, "G:\", TargetDriveLetter & ":\")  
        Tp = Replace(Tp, "M:\", TargetDriveLetter & ":\")  
        Tp = Replace(Tp, "N:\", TargetDriveLetter & ":\")  

        oSC.WorkingDirectory  = Wd         ' Update the shortcut's 
        oSC.TargetPath        = Tp         ' WorkingDirectory and TargetPath
        oSC.Description       = "-> " & Tp ' Indicate we're modifying the shortcut
        oSC.WindowStyle       = 1          ' Default property window style

        ' Inform the user how we change the shortcut's references.
        ' A timed-out popup displays the adjusted new shortcut.
        '
        'WshShell.Popup "-> Ausführen in:" & vbCRLF & oSC.WorkingDirectory & vbCRLF & _
        '               "-> Ziel:" & vbCRLF & oSC.TargetPath, _
        '               TIMEOUT_1s, oFile.Name

        oSC.Save  ' Finally save the updated shortcut
    End If        ' Process next shortcut
Next              ' Done.

WshShell.Popup "-> Ausführen in:" & vbCRLF & oSC.WorkingDirectory & vbCRLF & _
               "-> Ziel:" & vbCRLF & oSC.TargetPath

' WORKAROUND: Remove the subst (see the detailed explanation above).
'
ret = WshShell.Run ("cmd /c subst " & TargetDriveLetter & ": /d", 0, TRUE)    

If Err.Number = 0 Then                                             ' If error-free
    WshShell.Popup "Verknüpfungen wurden geändert.", 2*TIMEOUT_1s  '  tell the user we're done
End If

On Error Goto 0   ' Turn error handling off  

