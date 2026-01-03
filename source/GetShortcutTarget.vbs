REM ------------------------------------------------------------------------------- 
REM              -----------------------------------
REM             | VBScript: "GetShortcutTarget.vbs" |
REM              -----------------------------------
REM
REM  $_SCOPE  : Windows XP SP 3 (Language German)
REM                
REM  $_GENERAL: This VBscript copies the target file from a shortcut in to the 
REM             directory where the shortcut resides. Optionally, the
REM             copy of the target may be renamed.
REM             Usage:
REM              Copy the script named "GetShortcutTarget.vbs" to
REM              "C:\Dokumente und Einstellungen\UserName\SendTo".
REM              In Windows Explorer click the right mouse-button on the
REM              shortcut(s) or on the folder containing your shortcut(s),
REM              and, in the context menu choose "Senden an" 
REM              and select "GetShortcutTarget.vbs".
REM              This will run the script and process the shortcuts appropriately.
REM                
REM  $_NOTES  : V0.98
REM
REM  $_DATE   : 30.07.2013
REM  $_AUTHOR : HelmutAltmann
REM
REM  $_NOTES  : V1.00
REM
REM  $_DATE   : 09.12.2020
REM  $_AUTHOR : HelmutAltmann
REM ------------------------------------------------------------------------------- 
'#Option Explicit

Dim i, j, k, l, ret, oArgs, intButton, intMsgButton, strtst, NumFiles
Dim DestinationFolder, DestinationFileName
Const TIMEOUT_1s = 1    ' Hold Popup for 1s 


'------------------------------------------------
'       Progress Bar Initialization (IE)         |
'-------------------------------------------GetShortcutTarget.vbs-----
Dim ScreenWidth, ScreenHeight, Title_ProgressDisplay, ProgressBarWidth, ProgressBarHeight
Dim objIE                   ' objExplorer needed for Function "ProgressDisplay"

Call GetMonitorProperties() ' Determine the monitor's rendition and capabilities

Title_ProgressDisplay = "Progress-Bar"
ProgressBarWidth  = 400
ProgressBarHeight = 150


l = 0
k = 0
intButton = vbAbort     ' Default is "No Files Processed."
Set WshShell = WScript.CreateObject("WScript.Shell")

Set oFSO  = CreateObject("Scripting.FileSystemObject")
Set oArgs = WScript.Arguments

On Error Resume Next    ' Turn error handling on (may also hide some syntax errors) 

Set oFile = oFSO.GetFile(oArgs(0))          ' Check If 1st arg from commandline is a file

'#'----------------------------------------------------------------------
'#WshShell.Popup  "TESTPOINT1" & vbCrLf & "Err.Number: " & Err.Number   '|
'#'----------------------------------------------------------------------
If Err.Number <> 0 Then                     ' It's not a file If Error=53, it may be a folder
    Err.Clear                               ' Clean error object

    Set oFile = Nothing                     ' Prepare for re-usage after error
    Set oFolder = oFSO.GetFolder(oArgs(0))  ' Assign 1st commandline argument, that obviously
                                            '  is the folder's path passed to 'sendto' 

    If Err.Number <> 0 Then                 ' Here we should abort on error
        WScript.Echo "Fehler:"  & vbTab & Err.Number & " (" & Hex(Err.Number) & "h)" & vbCrLf & _
                     "Quelle:"  & vbTab & Err.Source & vbCrLf & _
                     "Ursache:" & vbTab & Err.Description
        Err.Clear                           ' Reset (clean) the Err object
        WScript.Quit                        ' Abort to Operation System.                                       
    End If

'-----------------------------------------------
' Assuming the commandline argument is a Folder |
'-----------------------------------------------
    intMsgButton = MsgBox("Do you want to show additional information when copying?", _
                           vbYesNoCancel+vbSystemModal, _
                           " - Processing Directory-Shortcuts - ")

    If intMsgButton = vbNo Then             ' Last chance to stop copying all files into folder
        intMsgButton = MsgBox("Copying files ...", _
                              vbOKCancel+vbQuestion+vbSystemModal, _
                              " - Processing Directory-Shortcuts - ")
    End If

    ProgressDisplay "Open",""               ' Get initial display configuration (needed for progress bar)

'#  NumFiles = oFolder.Files.Count          ' Number of (all!) files residing in folder
    NumFiles = 0                            ' Actually we should check and count only files with ".LNK" suffix!
    For Each oFile In oFolder.Files         ' Consult all files, but count only the shortcuts
        If UCase(right(oFile.Name,4)) = ".LNK" Then
            NumFiles = NumFiles + 1
        End If
    Next

    j = 2   ' Forces a Popup on exit
    For Each oFile In oFolder.Files         ' Consult all files, but process the shortcuts only

        If UCase(right(oFile.Name,4)) = ".LNK" Then

            ' Connect current shortcut and target file and perform the loop.
            ' The special targetname is derived from the shortcut's name.
            ' 
            Set oSC = WshShell.CreateShortcut(oFile.Path)           ' Connect next shortcut
            Set oTargetFile = oFSO.GetFile(oSC.Targetpath)          '  with corresponding targetpath

            ' Build destination directory (where the shortcut resides)
            '  that is the folder's path passed via 'sendto'
            '
            Call ParseShortcutFilename(oFile.Name, oTargetFile.Name)' Construct the final filename

            If intButton = vbCancel Then Exit For   ' Exit - "Cancel" was pressed when running

            ' Copy the renamed target file as [DestinationFolder\DestinationFileName]
            '  into the destination directory
            '
            oTargetFile.Copy DestinationFolder & DestinationFileName, TRUE
            k = k + 1                       ' Count the files processed     

            Select Case intMsgButton        ' Check the behaviour requested when invoked
                Case vbYES
                    Call DisplayFileInfo()  ' Show additional info box (returns [intButton])
                Case vbCancel               ' Force cancel by means of [intButton]
                    intButton = vbCancel    ' Popup: Cancelled - Stop and exit 
                Case vbNo
                    intButton = vbTrue      ' Popup: Done
                Case vbOK
                    intButton = vbTrue      ' Popup: Done
                    Call ShowProgress(k, NumFiles)  ' Display a progress bar
            End Select  ' Case intMsgButton

        End If

    next    ' Get next shortcut target in directory folder  

    ProgressDisplay "Close",""              ' Close Internet Explorer instance (Progress Bar)


'---------------------------------------------------
' Assuming the commandline arguments convey file(s) |
'---------------------------------------------------

Else    ' Else [err.number]

    j = oArgs.Count
    For i = 0 To oArgs.Count - 1
        Set oFile = oFSO.GetFile(oArgs(i))                  ' Get next commandline argument

        If Err.Number <> 0 Then                             ' Here we should abort on error
            WScript.Echo "Fehler:" & vbTab & Err.Number & " (" & Hex(Err.Number) & "h)" & vbCrLf & _
                         "Quelle:" & vbTab & Err.Source & vbCrLf & _
                         "Ursache:" & vbTab & Err.Description
            Err.Clear                                       ' Reset (clean) the Err object
            WScript.Quit                                    ' Abort to Operating System.                                       
        End If

        If UCase(Right(oFile.Name, 4)) = ".LNK" Then        ' Only shortcuts will be processed

            ' Connect current shortcut and target file and perform the loop.
            ' The special targetname is derived from the shortcut's name.
            ' 
            Set oSC = WshShell.CreateShortcut(oFile.Path)   ' Connect next shortcut
            Set oTargetFile = oFSO.GetFile(oSC.Targetpath)  '  with target   

            ' Build destination directory (where the shortcut resides)
            '  that is the file's path passed via 'sendto'
            '
            Call ParseShortcutFilename(oFile.Name, oTargetFile.Name)    ' Construct the final filename
            Call DisplayFileInfo()                                      ' Information Popup for the user

            If intButton = vbCancel Then Exit For           ' Exit - "Cancel" was pressed when running

            ' Copy the renamed target file as [DestinationFolder\DestinationFileName]
            '  into the destination directory
            '
            oTargetFile.Copy DestinationFolder & DestinationFileName, TRUE
            k = k + 1               ' Count the files processed     
        End If  

    Next    ' One shortcut target processed. Get next shortcut target                                                   
   
End If  ' End If [err.number]
'
'--------------------------------------------------------------------------

'---------------------------------
' Tell the user about termination |
'---------------------------------

If (j >= 2) OR (intButton = vbAbort) Then   ' Only show popup If more than one file
    Select Case intButton                   '  or If file is no legal shortcut
        Case vbTrue
            WshShell.Popup "Done. " & k & " files copied.",,, vbSystemModal
        Case vbCancel
            WshShell.Popup "Cancelled! " & k & " files copied.",,, vbSystemModal
        Case vbAbort
            WshShell.Popup "No files processed.",,, vbSystemModal
    End Select
End If

On Error Goto 0 ' Turn error handling off  

Set WshShell = Nothing

'# -------------------------------------------------------------------------------------
'#                          ParseShortcutFilename                                       '|
'#                                                                                      '|
'#          ' Build new DestinationFileName from the ShortcutName,                      '|
'#          '  If encountered specially named MP3 Tonband Shortcuts.                    '|
'#          '  Example of an MP3 Tonband shortcut:                                      '|
'#          '   "+A11_47 The Move - Night of Fear.mp3.LNK"                              '|
'#          '   After retrieval the copy of the target will named as derived            '|
'#          '   from the correlated shortcut's name and renamed to:                     '|
'#          '   "The Move - Night of Fear.mp3"                                          '|
'#          '                                                                           '|
'#          If UCase(right(oFile.Name,8)) = ".LNK.LNK" Then                             '|
'#              DestinationFileName = Mid(oFile.Name, 9, Len(oFile.Name)-16)            '|
'#          Else                                                                        '|
'#              DestinationFileName = oTargetFile.Name ' The unchanged TargetFileName   '| 
'#          End If                                                                      '| 
'#                                                                                      '| 
'#          ' Copy the target file into the destination directory                       '| 
'# -------------------------------------------------------------------------------------
Sub ParseShortcutFilename(Filename, TargetFilename)

    strtst = InstrRev(UCase(oFile.Name),".WAVMP3",-1,1) ' Handle special .wavmp3 shortcut
    If strtst > 1 Then
        DestinationFileName = Mid(Replace(oFile.Name, "+","_"), 1, strtst-1)'Keep "_A12_" prefix
    End If

    DestinationFolder = Left(oFile.path, Len(oFile.Path) - Len(oFile.name))
     
    If UCase(right(Filename,8)) = ".LNK.LNK" Then

        If Left(Filename,8) = "+WavMp3 " Then       ' Check special MP3 shortcut
            DestinationFileName = Mid(Filename, 9, Len(Filename)-8-8)

        ElseIf Left(Filename,1) = "+" AND (Instr(Filename,"_") = 5 AND Instr(Filename," ")) = 8 Then
'#          DestinationFileName = Mid(Filename, 6, Len(Filename)-5-8)   ' Tonband: eliminate "+A12_" prefix
            DestinationFileName = Mid(Filename, 2, Len(Filename)-1-8)   ' Tonband: Keep "A12_" prefix

        ElseIf Left(Filename,1) = "+" Then

            If UCase(InstrRev(1,Filename,".WAVMP3",1)) > 1 Then
                DestinationFileName = Mid(Filename, 5, Len(Filename)-15-4)
            Else        
                DestinationFileName = Mid(Filename, 2, Len(Filename)-8-1)
            End If

        End If

    ElseIf Left(Filename,16) = "Verknüpfung mit " Then                  ' Windows XP style       
        DestinationFileName = Mid(Filename, 17, Len(Filename)-16-4)

    ElseIf Right(Filename,14) = " - Verknüpfung" Then                   ' Windows 10 style       
        DestinationFileName = Mid(Filename, 14, Len(Filename)-14-4)

    ElseIf strtst < 1 Then
        DestinationFileName = TargetFilename    ' The unchanged TargetFileName
                                                '  otherwise DestinationFilename set previously
    End If                                      

End Sub ' ParseShortcutFilename

'# -------------------------------------------------------------------------------------
'#                          DisplayFileInfo                                             '|
'#                                                                                      '|
'#          ' Information Popup for the user (vbSystemModal = Always on top)            '|                                                                          '|
'#          ' Returns the [intButton] user response                                     '|
'# -------------------------------------------------------------------------------------
Sub DisplayFileInfo()

    intButton = WshShell.Popup ("Copy" & vbCrLf & oTargetFile.Path & vbCrLf & vbCrLf & _        
                                "to" & vbCrLf & DestinationFolder & DestinationFileName, _                  
                                TIMEOUT_1s,, vbOKCancel+vbExclamation+vbSystemModal)        

End Sub ' DisplayFileInfo

'# -------------------------------------------------------------------------------------
'#                          Progress-Bar                                                '|
'#                                                                                      '|
'#          ' Implementation  by means of the Internet Explorer,                        '|
'#          ' since Microsoft's vbScript won't support message boxes without prompt     '|
'# -------------------------------------------------------------------------------------
Sub ShowProgress(Progress0to100, MaxCount): Dim Text, n: n = (ProgressBarWidth - 2*19-1)
    Text = "<p align=""center"">Progress " & CStr(Round((Progress0to100/MaxCount)*100,0)) & " %</p>" & _
            "<table border=""0"" cellpadding=""0"" cellspacing=""0""><tr><td width=""" & _
            CStr(n*(Progress0to100)/MaxCount) & _
            """ height=""15"" bgcolor=""#0000FF"">&nbsp;</td></tr></table>"
    ProgressDisplay "Display",Text      ' Alternatively "Call ProgressDisplay("Display",Text)"
    WScript.Sleep 1
End Sub ' ShowProgress

Sub ProgressDisplay (Mode, AnyText): Dim String1, String2, colItems, objItem
    ' Mode = Open, Display, Close
    ' AnyText only used in Display-Mode
    Mode = UCase(Left(Mode,1)) & LCase(Right(Mode,Len(Mode)-1))     ' Normalize Mode text-string
    Select Case Mode
        Case "Open"
            Set objIE = CreateObject("InternetExplorer.Application")' Using routines of IE
            With objIE                                              ' IE specific variables
                .Navigate "about:blank"                             ' Initialize the IE progress-bar window
                .ToolBar = False: .StatusBar = False
                .Width = ProgressBarWidth: .Height = ProgressBarHeight
                .Left = (ScreenWidth - ProgressBarWidth) \ 2
                .Top = (ScreenHeight - ProgressBarHeight) \ 2
                .Visible = True
                With .Document
                    .title = Title_ProgressDisplay
                    .ParentWindow.focus()
                    With .Body.Style
                        .backgroundcolor = "#F0F7FE"
                        .color = "#0060FF"
                        .Font = "11pt 'Calibri'"
                    End With
                End With
                While .Busy: Wend
            End With

        Case "Display"                  ' Display the progress-bar Popup window (IE)
            On Error Resume Next        ' For clicking away the bar while running
            If Err.Number = 0 Then
                With objIE.Document             ' Display the (updated) progress-bar
                    WScript.Sleep 1             ' Slow (=200) / Fast (=1) progress-bar
                    .Body.InnerHTML = AnyText   ' This is the "HTML-Page" as proggress-bar
                    .ParentWindow.focus()       ' Pops up the IE window on top
                End With
            End If

        Case "Close"                    ' Exit the IE application
            WScript.Sleep 100
            objIE.Quit
    End Select
End Sub ' ProgressDisplay

Sub GetMonitorProperties
    Dim strComputer, objWMIService, objItem, colItems, VMD: strComputer = "."

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController")

    For Each objItem In colItems: VMD = objItem.VideoModeDescription: Next

    ' VMD = 1280 x 1024 x 4294967296 Farben
    VMD = Split(VMD, " x "): ScreenWidth = Eval(VMD(0)): ScreenHeight = Eval(VMD(1))
End Sub ' GetMonitorProperties

'# -------------------------------------------------------------------------------------
'#                                                                                      '|
'#                  HTML principle of the Progress-Bar in use                           '|
'#                                                                                      '|
'# -------------------------------------------------------------------------------------

'#<!DOCTYPE html>
'#<html>
'#
'#<head>
'#  <title>Progress-Bar</title> <!-- Visible in IE Window-Tab -->
'#</head>
'#
'#
'#<style>
'#
'#body {
'#  background-color: #F0F7FE;
'#  color: #0060FF;
'#  font-size: 11pt;  font-family: Calibri;
'#}
'#
'#p {
'#  text-align: center;
'#}
'#
'#table, td {
'#  border: 0;
'#  cellpadding: 0;
'#  cellspacing: 0;
'#  background-color: #0060FF;
'#  color: #0060FF;
'#  visibility: visible;
'#}
'#
'#table td {
'#  width: 10%;
'#}
'#
'#div.relative {
'#  position: absolute;
'#  Top: 400px;
'#  left: 400px;
'#  width: 20%;
'#  height: 15%;
'#  border: 3px solid #73AD21;
'#  }
'#
'#</style>
'#</head>
'#
'#<body>
'#
'#<div class="relative">
'#
'#<p>Progress 100 %</p>
'#
'#<table>
'#  <tr>
'#    <td> &nbsp; </td>
'#  </tr>
'#</table>
'#
'#</div>
'#
'#</body>
'#</html>

'# -------------------------------------------------------------------------------------
'#          Implemtation here: Partly without using CSS-Style
'# -------------------------------------------------------------------------------------
'#<style>
'#body {
'#  background-color: #F0F7FE;
'#  color: #0060FF;
'#  font-size: 11pt;  font-family: Calibri;
'#}
'#
'#table.td-width td {width: 400}
'#Table.td-height td {height: 150}
'#
'#
'#div.relative {
'#  position: absolute;
'#  Top: 400px;     
'#  left: 400px;        
'#  width: 20%;
'#  height: 15%;
'#  border: 3px solid #73AD21;
'#  }
'#
'#</style>
'#
'#<body>
'#
'#<div class="relative">
'#
'#<p align="center">Progress 100 %</p>
'#
'#<table class="td-width" class="td-height" border="0" cellpadding="0" cellspacing="0">
'#  <tr>
'#    <td width="214" height="25" bgcolor="#0000FF"> &nbsp; </td>
'#  </tr>
'#</table>
'#</div>
'#
'#</body>

'# -------------------------------------------------------------------------------------
'#          Implemtation with inbuild Browser Functions (see vbScript)
'# -------------------------------------------------------------------------------------
'#<!DOCTYPE html>
'#<html>
'#
'#<head>
'#  <title>Progress-Bar</title> <!-- Visible in IE Window-Tab -->
'#</head>
'#
'#
'#<body>
'#<p>Click the button to open an about:blank page in a new browser window.</p>
'#<button onclick="myFunction()">Try it</button>
'#
'#<script>
'#function myFunction() {
'#  var myWindow = window.open("", "", "width=400, height=150, left=580, top=545");
'#
'#<!-- firefox ignores these statements, must be given in "window.open()"
'#  myWindow.navigate = "about.blank"
'#  myWindow.toolbar = false
'#  myWindow.statusbar = false
'#  myWindow.width = 400
'#  myWindow.height = 150
'#  myWindow.left = 580
'#  myWindow.top = 545
'#  myWindow.visible = true
'#-->
'#
'#<!--   myWindow.document.ParentWindow.focus() -->
'#
'#  myWindow.document.title = "Progress-Bar";
'#
'#  myWindow.document.body.style.backgroundColor= "#F0F7FE";
'#  myWindow.document.body.style.color          = "#0060FF";
'#  myWindow.document.body.style.font           = "11pt 'Calibri'";
'#
'#  myWindow.document.body.innerHTML = "<p align='center'>Progress 100 %</p> <table border='0'><tr><td width='214' height='15' bgcolor='#0000FF'>&nbsp;</td></tr></table>" 
'#}
'#</script>
'#
'#</body>
'#</html>
