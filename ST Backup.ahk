#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent 
#SingleInstance, ignore 
#NoTrayIcon

#Include Class_CtlColors.ahk

sPath := ""
sDest := ""
_sPath := ""
_sDest := ""
sMainLogName := "\stbackup_log.txt"
sMainLogPath := ""
sBackupLogPath := ""
sCustomDest := ""
sExts := ""
iBackupCount := 10
tInterval := 300000 ; 5 min
toggle := 0
sCurrentTime :=""
bCopyallExts:=false
bRecursive:=false
bZipBackup := 0
errIcon := 16
infoIcon := 64
curVersion:=1.1
myName:="Simple Timed Backup"
myTitle:=myName . " " . curVersion
bLogVarEnabled := False
winHeight_LogHide := 228
winHeight_LogShow := 328
iMaxLogSize := 500 ;kb

;DllCall("AllocConsole")
;WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")


_font:="Tahoma"

IsEmpty(Dir){
   Loop %Dir%\*.*, 0, 1
      return 0
   return 1
}

Zip(sDir, sZip)
{
    If Not FileExist(sZip)
    {
        Header1 := "PK" . Chr(5) . Chr(6)
        VarSetCapacity(Header2, 18, 0)
        file := FileOpen(sZip,"w")
        file.Write(Header1)
        file.RawWrite(Header2,18)
        file.close()
    }
    psh := ComObjCreate( "Shell.Application" )
    pzip := psh.Namespace( sZip )
    pzip.CopyHere( sDir, 4|16 )
    Loop {
        sleep 100
        zippedItems := pzip.Items().count
        ;ToolTip Zipping in progress..
    } Until zippedItems=1 ;because sDir is just one file or folder
    ;ToolTip
}



Unz(sZip, sUnz)
{
    fso := ComObjCreate("Scripting.FileSystemObject")
    If Not fso.FolderExists(sUnz)
    fso.CreateFolder(sUnz)
    psh  := ComObjCreate("Shell.Application")
    zippedItems := psh.Namespace( sZip ).items().count
    psh.Namespace( sUnz ).CopyHere( psh.Namespace( sZip ).items, 4|16 )
    Loop {
        sleep 100
        unzippedItems := psh.Namespace( sUnz ).items().count
        ;ToolTip Unzipping in progress..
        IfEqual,zippedItems,%unzippedItems%
            break
    }
    ;ToolTip
}

psZip(inPath,outPath)
{
    RunWait PowerShell.exe -Command Compress-Archive -LiteralPath '%inPath%' -CompressionLevel Optimal -DestinationPath '%outPath%',, Hide UseErrorLevel
    Return ErrorLevel
}
psUnzip(inPath,outPath)
{
    RunWait PowerShell.exe -Command Expand-Archive -LiteralPath '%inPath%' -DestinationPath '%outPath%',, Hide UseErrorLevel
    Return ErrorLevel
}

zipBackup(sPath)
{
    if InStr(FileExist(sPath), "D") {
        if IsEmpty(sPath) {
            return
        }
    }
    else {
        return
    }
    SplitPath, sPath, sName, sParent
    FileDelete, %sParent%\%sName%.zip
    FileDelete, %sPath%\%sName%.zip
    sOut := sParent "\" sName ".zip"
    Run PowerShell.exe -Command (Compress-Archive -LiteralPath '%sPath%' -CompressionLevel Optimal -DestinationPath '%sOut%'); if ($?) { (Remove-Item -force '%sPath%' -recurse -Confirm:$False); (New-Item -force -Path '%sPath%'  -ItemType Directory); (move-Item '%sOut%' '%sPath%' -force); },, Hide UseErrorLevel
    if (ErrorLevel = "ERROR")
    {
        Zip(sPath , sParent "\" sName ".zip")      
        FileRemoveDir, %sPath%, 1
        FileCreateDir, %sPath%
        FileMove, %sParent%\%sName%.zip,%sPath%,1
    }
    Return
}

trimExts(ByRef sExtensions)
{
    StringReplace, sExtensions, sExtensions,`n,,All
    StringReplace, sExtensions, sExtensions,%A_SPACE%,, All
    StringReplace, sExtensions, sExtensions,%A_Tab%,, All
    StringReplace, sExtensions, sExtensions,.,, All
    StringReplace, sExtensions, sExtensions,/,, All
    StringReplace, sExtensions, sExtensions,\,, All
    StringReplace, sExtensions, sExtensions,:,, All
    StringReplace, sExtensions, sExtensions,|,, All
    StringReplace, sExtensions, sExtensions,",, All ;"a comment to fix notpad++ Syntax Highlighting
    StringReplace, sExtensions, sExtensions,<,, All
    StringReplace, sExtensions, sExtensions,>,, All
    StringReplace, sExtensions, sExtensions,`,,, All
    if(sExtensions ="")
    {
        sExtensions := "*;"
    }
}

logErrors(sExt,sBackupPath,errCount,bSilent:=true)
{
    global sPath
    global sDest
    global mainStatusBarVar_TT
    Global sBackupLogPath
    Global sMainLogName
    Global iMaxLogSize
    Global sMainLogPath
    FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
    FormatTime, curTime, %a_now% T12, [HH:mm]
    if (errCount < 0)
    {
        strLog := "Warning: No file copied. Type=*." . sExt
        logEditAdd(shrinkString(strLog,73,"r"))
    }
    else if (errCount = 0)
    {
        FormatTime, sCurrentTime ,  dddd MMMM d yyyy HH:mm:ss T12
        if FileExist(sMainLogPath)
        {
            FileGetSize, logsizekb, %sMainLogPath%, K
            if (logsizekb>iMaxLogSize)
            {
                FileDelete, %sMainLogPath%
                FileAppend ,%sNow% backup started..., %sMainLogPath%
                FileAppend ,`n%sNow% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            } 
            else  
            {
                FileAppend ,`n%sNow% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            }
        } 
        else  
        {
            FileAppend ,%sNow% backup started..., %sMainLogPath%
            FileAppend ,`n%sNow% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
        }           
        mainStatusBarVar_TT := sBackupPath
        strLog := shrinkString("*." . sExt " Backup: """ . trimPath(sBackupPath) . """",73,"r")
        SB_SetText(A_Tab  . curTime . " " . strLog,1,1)
        logEditAdd(strLog)
        if (FileExist(sBackupLogPath)) 
        {
            FileAppend ,`n*.%sExt% Backup: %sCurrentTime%,%sBackupLogPath%
        }
        else
        {
            FileAppend ,*.%sExt% Backup: %sCurrentTime%,%sBackupLogPath%
        } 
    } 
    else 
    {
        if FileExist(sMainLogPath)
        {
            FileAppend ,`n%sNow% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            FileAppend ,`n%sNow% can`t copy %errCount% file(s)!
        }
        else
        {
            FileAppend ,%sNow% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            FileAppend ,`n%sNow% can`t copy %errCount% file(s)!
        }
        strLog := shrinkString("Error: Cannot copy " . errCount . " file(s) to destination. Type=*." . sExt, 73, r)
        logEditAdd(strLog)
        if (!bSilent) 
        {
            SplashTextOff
            msgBox,% errIcon,, Cannot copy some files!
            Gosub, ExitSub
        }
    }
    Return
}

trimPath(strPath)
{   
    if (strPath="")
    {
        return strPath
    }
    break:=false
    while(break=false)
    {
        break:=true
        if (SubStr(strPath,StrLen(strPath)) = "\") 
        {
            strPath:=substr(strPath,1,StrLen(strPath)-1) 
            break:=false
        }
        if (SubStr(strPath,StrLen(strPath)) = " ") 
        {
            strPath:=substr(strPath,1,StrLen(strPath)-1)
            break:=false
        }
        if (SubStr(strPath,1,1) = " ") 
        {
            strPath:=substr(strPath,2)
            break:=false
        }
    }
    return strPath
}

shrinkString(strIn,maxLength,side)
{
    StringLower,side,side
    if (StrLen(strIn) > maxLength) {
        if (side="l") {
            StringTrimLeft,strIn,strIn,StrLen(strIn) - maxLength
            strIn := "..." . strIn
        }
        else if (side="m") {
            tempStr:=SubStr(strIn,StrLen(strIn)//2)
            StringTrimLeft,tempStr,tempStr,StrLen(strIn) - maxLength
            if tempStr=
                tempStr:=SubStr(strIn,0) 
            strIn := SubStr(strIn,1, min(StrLen(strIn)//2-1,maxLength-1)) . "..." . tempStr
        }
        else if (side="r") {
             StringTrimRight,strIn,strIn,StrLen(strIn) - maxLength
             strIn := strIn . "..."
             }
       }
        return strIn
}

bIsParentPath(sParentPath,sChildPath)
{
    If (StrReplace(sChildPath,sParentPath)=sChildPath)
    {
        return false
    }
    return true
}

checkNum(ByRef Variable, ControlID, MinVal, MaxVal)
{
	hasSpaces:=RegExMatch(Variable,"[\s]")		; Check if contains spaces
    numVar:=RegExReplace(Variable,"[^0-9.]+")	; Remove all non numerics or .
	ControlGet,inPos, CurrentCol,, Edit1,A		; Get input position of the edit box.
	StringSplit,splitNum,numVar,.				; Allow only one ".", the leftmost is preserved.
	if (splitNum0>2)							; This can be higher than three if user pastes in something with more than one dot.
	{
		numVar:=splitNum1 . "."
		Loop, % splitNum0-1
		{
			ind:=A_Index+1
			numVar.=splitNum%ind%	 
		}	
	}
    if (numVar > MaxVal) {
        numVar := MaxVal        
    }
    else if (numVar < MinVal) {
        numVar := MinVal        
    }
	if (Variable==numVar && !hasSpaces) 		; If nothing changed and no spaces present, return
		return
    Variable := numVar
	GuiControl,,%ControlID%,%numVar%		; Set text
	PostMessage,0x00B1,inPos-2,inPos-2,Edit1,A 	; Move input caret to correct position, EM_SETSEL:=0x00B1
}

StrReplaceVar(strIn)
{
    ;replace the substring between two ? with content of a variable with the same name
    index:=1
    break:=false
    while(break=false) {
        break:=true
        ChrIndex:=InStr(strIn,"?",,index,1)
        if (ChrIndex > 0){
            nextIndex:=InStr(strIn,"?",,ChrIndex+1,1)
            if (nextIndex>0) {
                strVar:=SubStr(strIn,ChrIndex + 1,(nextIndex-ChrIndex)-1)
                strVar := %strVar%
                strIn:= StrReplace(SubStr(strIn,1,ChrIndex), "?") . strVar . SubStr(strIn,nextIndex+1)
                index := nextIndex+1
                if (index < strLen(strIn) - 1) {
                    break:=false
                }
            }
        }
    }
    return strIn
}

CopyFiles(sExt,sDestFolder,bRecursive:=false)   
{
    global sPath
    global sDest
    sPath:= trimPath(sPath)
    sDest:= trimPath(sDest)
    sDestFolder:= trimPath(sDestFolder)
    if IsEmpty(sPath)
    {
        return -1
    }
    FileCreateDir, %sDestFolder%
    ErrorCount:=ErrorLevel
    if(ErrorCount<>0)
    {
        return
    }
    ErrorCount := 0
    if (bRecursive=false) {
        FileCopy, %sPath%\*.%sExt%, %sDestFolder%\, 1
        ErrorCount := ErrorLevel
        If (ErrorCount > 0)
        {
            Return ErrorCount
        }
    } else {
        Loop Files, %sPath%\*.%sExt%, R  ; Recurse into subfolders.
        {
            if (A_LoopFileDir=sPath)
            {
                FileCopy, %A_LoopFileFullPath%, %sDestFolder%\, 1
            }
            else
            {
                sDestFileLongPath := StrReplace(A_LoopFileLongPath,sPath,sDestFolder)
                SplitPath, sDestFileLongPath,, sDestFileDir
                If (bIsParentPath(sDest,A_LoopFileLongPath))
                {
                    ;The file is a backup that's inside the source folder
                    Continue
                }
                if (InStr(FileExist(sDestFileDir),"D")=0)
                {
                    FileCreateDir, %sDestFileDir%                               
                }
                FileCopy, %A_LoopFileFullPath%, %sDestFileDir%\, 1                          
            }
            ErrorCount := ErrorLevel
            If (ErrorCount > 0)
            {
                Return ErrorCount
            }
        }
    }
    if IsEmpty(sDestFolder)
    {
        return -1
    }
    return 0
}

if (FileExist("STB_settings.ini"))
{
    IniRead, _sPath, STB_settings.ini, Paths, Files Location 
    IniRead, _sDest, STB_settings.ini, Paths, Backups Location
    IniRead, sCustomDest, STB_settings.ini, History, Last Manual Backup Location
    IniRead, tInterval, STB_settings.ini, Option, Backup Interval , 300000 
    IniRead, iBackupCount, STB_settings.ini, Option, Backups Count , 10
    IniRead, iMaxLogSize, STB_settings.ini, Option, Max Log Size , 500
    IniRead, iBkupNum, STB_settings.ini, History, Next Backup Number, 1
    IniRead, sExts, STB_settings.ini, Option , Extensions, "*;"
    IniRead, bZipBackup, STB_settings.ini, Option, Zip Backups , 0
    IniRead, bRecursive, STB_settings.ini, Option, Recursive , 0
    IniRead, bLogVarEnabled, STB_settings.ini, Option, Show log , 0
    sPath := StrReplaceVar(trimPath(_sPath))
    sDest := StrReplaceVar(trimPath(_sDest))
    if (sPath<>"") 
    {
        if (sPath=sDest) 
        {
            sDest .= "\ST_Backups"
            IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location
            
        }        
    }
    sMainLogPath := sDest . sMainLogName
    sCustomDest := trimPath(sCustomDest)
    if (sCustomDest<>"" and sCustomDest=sPath) {
        sCustomDest .="\ST_Backups"
        IniWrite, %sCustomDest%, STB_settings.ini, History, Last Manual Backup Location
    }
}else  {
    sExts:= "*;"
    IniWrite, %sPath%, STB_settings.ini, Paths, Files Location
    IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location
    IniWrite, %sCustomDest%, STB_settings.ini, History, Last Manual Backup Location
    IniWrite, %tInterval%, STB_settings.ini, Option, Backup Interval 
    IniWrite, %iBackupCount%, STB_settings.ini, Option, Backups Count
    IniWrite, %iMaxLogSize%, STB_settings.ini, Option, Max Log Size
    IniWrite, %iBkupNum%, STB_settings.ini, History, Next Backup Number
    IniWrite, %sExts%, STB_settings.ini, Option, Extensions
    IniWrite, %bZipBackup%, STB_settings.ini, Option, Zip Backups
    IniWrite, %bRecursive%, STB_settings.ini, Option, Recursive
    IniWrite, %bLogVarEnabled%, STB_settings.ini, Option, Show log
}

Hotkey, ^!x, ExitSub
OnExit, ExitSub

Gui +LastFound
GUI, -ToolWindow
Gui, +CAPTION
GUI, -MaximizeBox
GUI, Margin,11,
Gui,Font, normal s8, %_font%
Gui,Add,Text,x9 y12 w85 %black% left, Files to Backup
Gui,Add,Text,x9 y47 w85 %black% left, Backups Location
Gui,Add,Edit,x95 y10 w332 h30 HwndHSLedit r1 %black% vSLedit gRevertSLeditColor,
GuiControl,, SLedit, %sPath%
Gui,Add,Edit,x95 y45 w332 h30 HwndHBLedit r1 %black% vBLedit gRevertBLeditColor,
Gui,Font, s8 normal, %_font%
GuiControl,, BLedit, %sDest%
Gui,Add,Button,x437 y9 w30 h23 r1 center vSPvar gSPbtn,...
Gui,Add,Button,x437 y44 w30 h23 r1 center vBPvar gBPbtn,...
Gui,Add, GroupBox, x5 y80 w310 h101,
Gui,Add,Text,x10 y97 w80 h13 left  %black%  ,Backup every
Gui,Add,Edit,x85 y95 w70 h18 %black% number vBIedit gBIedit
mInterval := (tInterval/60000)
Gui,Add,UpDown, 0x20  Range1-720 ,%mInterval%,vBIud
Gui,Add,Text,x10 y116 w80 h13 %black% left ,Backup count

Gui,Add,Edit,x85 y114 w70 h18 %black% Number vBCedit gBCedit
Gui,Add,UpDown, 0x20  Range1-100000,%iBackupCount%,vBCud

Gui,Add, GroupBox, x10 y134 w150 h41, Backup these file types
Gui,Add,Edit,x15 y150 w140 h20 %black% Lowercase vextsediVar gextsEdit,%sExts%

Gui,Add,Checkbox,x166 y97 w140 h20 %black% -Wrap  vZipBackupvar gZipBackupcbx,Zip backups

if (bZipBackup = 1)
{
    GuiControl,, ZipBackupvar, 1    
}
else
{
    GuiControl,, ZipBackupvar, 0
}

Gui,Add,Checkbox,x166 y119 w140 h20 %black% -Wrap  vRecursiveVar gRecursivecbx,Recursive

if (bRecursive = 1)
{
    GuiControl,, RecursiveVar, 1    
}
else
{
    GuiControl,, RecursiveVar, 0
}

Gui,Add,Button, x165 y139 w147 h35 center +Disabled  vBKvar gBKbtn , Manual Backup...

Gui,Add,Button,x320 y92 w147 h35 +Disabled vDEvar gDEbtn,Deactivate
Gui,Add,Button,x320 y139 w147 h35 center vACvar gACbtn,Activate
Gui,Font, s8 normal, %_font%
Gui,Add, Checkbox, x9 y184 vShowLogcbx gToggleLogcbx, Show Log
Gui,Add,Text,x166 y184 w80 h13 left  %black%  ,Max Log Size
Gui,Add,Edit,x242 y184 w70 h18 number %black% vLSedit gLSedit
Gui,Add,UpDown, 0x20  Range10-10000,%iMaxLogSize%,vLSud

if sPath !=
    GuiControl, Enabled, BKvar

Gui,Font, s8 normal, %_font%
Gui, Add, StatusBar,gmainStatusBar vmainStatusBarVar,Ready
Gui,Font, s7 , Lucida Console
Gui,add, edit, x9 y212 w458 r9 left ReadOnly vLogEditVar gLogEdit
Gui,Font, s8 normal, %_font%

if (bLogVarEnabled=False)
{
    GuiControl,, ShowLogcbx, 0
    GuiControl, Hide, LogEditVar
    Gui,Show, h%winHeight_LogHide%,%myTitle%
}
else
{
    GuiControl,, ShowLogcbx, 1
    GuiControl, Show, LogEditVar
    Gui,Show, h%winHeight_LogShow%,%myTitle%
  
}

;Tooltips
SLedit_TT := "The source folder."
BLedit_TT := "The destination folder for storing backups."
SPvar_TT := "Change the source folder."
BPvar_TT := "Change the destination folder."
BCedit_TT := "How many backups should be created before overwriting previous backups."
ACvar_TT := "First, a backup will be created inside the ""backup_0"" folder.`nThen automated backups will be created at the selected interval."
DEvar_TT := "First, a backup will be created inside the ""backup_00"" folder.`nThen creating automated backups will be stopped."
extsediVar_TT := "Extensions are separated by `;`n* means any extension"
BKvar_TT := "Takes a manual backup inside the selected folder."
ZipBackupvar_TT := "Toggles the compression of backups."
RecursiveVar_TT := "Toggles backup for files in subfolders."
BIedit_TT := "The time between auto-updates in minutes."
EDbtnvar_TT := "Edit what file types to backup."
mainStatusBarVar_TT := ""
LogEditVar_TT := ""
LSedit_TT := "Max allowed size of the log file in KB."

OnMessage(0x200, "WM_MOUSEMOVE")
OnMessage(0x0203, "WM_LBUTTONDBLCLK")

Return

WM_MOUSEMOVE()
{
    static CurrControl, PrevControl, _TT  ; _TT is kept blank for use by the ToolTip command below.
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl and not InStr(CurrControl, " "))
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, DisplayToolTip, 1000
        PrevControl := CurrControl
    }
    return

    DisplayToolTip:
    SetTimer, DisplayToolTip, Off
    ToolTip % %CurrControl%_TT  ; The leading percent sign tell it to use an expression.
    SetTimer, RemoveToolTip, 3000
    return

    RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
    return
}

WM_LBUTTONDBLCLK(wParam, lParam)
{
    if(a_guicontrol = "SLedit") {
        GuiControlGet, strPath,, SLedit
        if InStr(FileExist(strPath),"D")
        {
            Run, Explorer /n`,/e`,%strPath%
        }
    }
    if(a_guicontrol = "BLedit") {
        GuiControlGet, strPath,, BLedit
        if InStr(FileExist(strPath),"D")
        {
            Run, Explorer /n`,/e`,%strPath%
        }
    }
    if(a_guicontrol = "LogEditVar") {
        GuiControlGet, strPath,, BLedit
        if InStr(FileExist(strPath),"D")
        {   logPath := strPath . "\stbackup_log.txt"
            if (FileExist(logPath))
            {
                SplitPath, logPath,,,,fName 
                Run,% "notepad.exe " . logPath
                If WinExist(fName)
                    WinActivate 
            }
        }

    }
    if(a_guicontrol = "mainStatusBarVar") {
        global mainStatusBarVar_TT
        if (mainStatusBarVar_TT<>"")
        {
            if InStr(FileExist(mainStatusBarVar_TT),"D")
            {
                Run, Explorer /n`,/e`,%mainStatusBarVar_TT%
            }
        }
    }    
    return
}

RevertSLeditColor:
{
    CtlColors.Change(HSLedit, "", "")
    Return
}

RevertBLeditColor:
{
    CtlColors.Change(HBLedit, "", "")
    Return
}

mainStatusBar:
{
    return
}

ToggleLogcbx:
{
    GuiControlGet, bLogVarEnabled, , ShowLogcbx
    If (bLogVarEnabled=False)
    {
      GuiControl, Hide, LogEditVar
      Gui,Show, h%winHeight_LogHide%,%myTitle%
    }
    Else
    {
      GuiControl, Show, LogEditVar
      Gui,Show, h%winHeight_LogShow%,%myTitle%
      
    }
    IniWrite, %bLogVarEnabled%, STB_settings.ini, Option, Show log
    return
}

logEdit:
{
    return
}

logEditAdd(strIn)
{
     GuiControlGet, sLogEditText ,, LogEditVar
     FormatTime, curTime, %a_now% T12, [HH:mm:ss]
     sLogEditText := curTime . " " . strIn . "`r`n" . sLogEditText
     GuiControl,text, LogEditVar, %sLogEditText%
    return
}

SPbtn:
{
    FileSelectFolder,OutputVar1 ,*%sPath% , 0, Files location
    if OutputVar1 =
    return
    GuiControl,, SLedit, %OutputVar1%
    GuiControl, Enabled, BKvar
    sPath := OutputVar1
    IniWrite, %sPath%, STB_settings.ini, Paths, Files Location
    Return
}

BPbtn:
{
    FileSelectFolder,OutputVar2 ,*%sDest% , 3, Backups location
    if OutputVar2 =
    return
    if (OutputVar2 = sDest)
        return
    FileCreateDir, %OutputVar2%\ST_Backups
    OutputVar2 .= "\ST_Backups"
    GuiControl,, BLedit, %OutputVar2%
    sDest := OutputVar2
    sMainLogPath := sDest . sMainLogName
    IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location
    Return
}

BIedit:
{
    GuiControlGet ,InVar,, BIedit
    checkNum(InVar, "BIedit", 1,720)
    tInterval := InVar*60000
    IniWrite, %tInterval%, STB_settings.ini, Option, Backup Interval 
    Return
}

BCedit:
{
    GuiControlGet ,InVar,, BCedit
    checkNum(InVar, "BCedit", 1,100000)
    iBackupCount := InVar
    IniWrite, %iBackupCount%, STB_settings.ini, Option, Backups Count
    Return
}

LSedit:
{
    GuiControlGet, InVar,, LSedit
    checkNum(InVar, "LSedit", 10,100000)
    iMaxLogSize := InVar
    IniWrite, %iMaxLogSize%, STB_settings.ini, Option, Max Log Size
    Return
}

extsEdit:
{
    Return
}
    
ZipBackupcbx:
{
    bZipBackup := !bZipBackup
    IniWrite, %bZipBackup%, STB_settings.ini, Option, Zip Backups
    Return
}

Recursivecbx:
{
    bRecursive := !bRecursive
    IniWrite, %bRecursive%, STB_settings.ini, Option, Recursive
    Return    
}
    
ACbtn:
{ 
    Gui, Submit , NoHide
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, tInterVal,, BIedit
    GuiControlGet, iBackupCount,, BCedit
    GuiControlGet, Extstring ,, extsediVar,
    trimExts(Extstring)
    sPath := trimPath(sPath)
    sDest := trimPath(sDest)
    sExts := Extstring
    GuiControl,,extsediVar, %sExts% 
    StringSplit, ExtArr, Extstring ,`;,
    sPVar :=InStr(FileExist(spath),"D")
    if(tInterval="" )
    {
        tInterval := 300000
        GuiControl, , BIud,%tInterval%
    }
    else if (iBackupCount="")
    {
        iBackupCount := 10
    }
    else if (sPVar=0)
    {
        CtlColors.Change(HSLedit, "FFC0C0", "")
        GuiControl,Focus, SLedit
        Return
    }
    Else if tInterval not between 1 and 720
    {
        SplashTextOff
        msgbox,% errIcon,, Your Backup Interval is not within the valid range: 1-720
        return
    }
    Else if iBackupCount not between 1 and 100
    {
        SplashTextOff
        msgbox,% errIcon,, Your Backup Count is not within the valid range: 1-100
        return
    }
    else  
    {
        if (sPath=sDest)
        {
            if (_sPath<>"" and (_sPath=_sDest) and (StrReplaceVar(_sPath)=sPath) and (StrReplaceVar(_sDest)=sDest)) 
            {
                _sDest.="\ST_Backups"
                sDest.="\ST_Backups"
                IniWrite, %_sDest%, STB_settings.ini, Paths, Backups Location
            }
            else {
                sDest.="\ST_Backups"
                IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location
            }
            GuiControl,, BLedit, %sDest%
            
        }
        if (sDest="")
        {
            CtlColors.Change(HBLedit, "FFC0C0", "")
            GuiControl,Focus, BLedit
            Return        
        }
        sDVar :=InStr(FileExist(sDest),"D")
        If (sDVar=0)
        {
            FileCreateDir, %sDest%
            erl:=ErrorLevel
            if(erl<>0)
            {
                SplashTextOff
                msgbox,% errIcon,, The path you entered could not be created: %sDest%
                return
            }
        }
        GuiControl,, BLedit, %sDest%
        GuiControl,, SLedit, %sPath%
        tInterval:=tInterval*60000
        GuiControl,Disable,ACvar
        GuiControl,Enable,DEvar
        GuiControl,Disable,SPvar
        GuiControl,Disable,BPvar
        GuiControl,Disable,BCedit
        GuiControl,Disable,BIedit
        GuiControl,Disable,SLedit
        GuiControl,Disable,BLedit
        GuiControl,Disabled,EDbtnvar
        GuiControl, Disabled, ZipBackupvar
        GuiControl, Disabled, RecursiveVar
        GuiControl, +ReadOnly, extsediVar
        SB_SetText(A_Tab  . "Auto backup started.",1,1)
        logEditAdd("Auto backup started.")
        sMainLogPath := sDest . sMainLogName
        FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
        if FileExist(sMainLogPath)
        {
            FileAppend ,`n%sNow% Auto backup started..., %sMainLogPath%
        }
        else
        {
            FileAppend ,%sNow% Auto backup started..., %sMainLogPath%
        }
        sBackupPath := sDest . "\Backup_0"
        sBackupLogPath := sBackupPath . "\stbackup_log.txt"        
        FileDelete, %sBackupLogPath%
        FileRemoveDir, %sBackupPath%, 1
        if(iBkupNum="")
        {
            iBkupNum := 1
        }else  {
            if iBkupNum not between 1 and %iBackupCount%
            iBkupNum := 1
        }
        sleep,2000
        bCopyallExts:=false
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index%="*")
            {
                bCopyallExts:=true
                Break
            }
        }
        if(bCopyallExts = false)
        {
            loop, %ExtArr0%
            {
                if(ExtArr%A_Index% <> "")
                {
                    tempExt:=ExtArr%A_Index%
                    ErrorCount := CopyFiles(tempExt,sBackupPath,bRecursive)
                    logErrors(tempExt, sBackupPath, ErrorCount)
                    If (ErrorCount > 0)
                    {
                        Return
                    }
                }
            }
        } else If ( bCopyallExts = True)
         {
            ErrorCount := CopyFiles("*",sBackupPath,bRecursive)
            logErrors("*", sBackupPath, ErrorCount)
            If (ErrorCount > 0)
            {
                Return
            }
        }       
        If (bZipBackup = 1) {
            zipBackup(sDest "\Backup_0")
        }         
        Gosub, ToggleBackup
    }
    Return
}

DEbtn:
{
    GuiControl,Disable,DEvar
    GuiControl,Enable,ACvar
    GuiControl,Enable,SPvar
    GuiControl,Enable,BPvar
    GuiControl,Enable,BCedit
    GuiControl,Enable,BIedit
    GuiControl,Enable,SLedit
    GuiControl,Enable,BLedit
    GuiControl,Enable,ZipBackupvar
    GuiControl,Enable,RecursiveVar
    GuiControl, -ReadOnly, extsediVar
    sBackupPath := sDest . "\Backup_00"
    sBackupLogPath := sBackupPath . "\stbackup_log.txt"
    FileDelete, %sBackupLogPath%
    FileRemoveDir, %sBackupPath%, 1
    if(bCopyallExts = false)
    {
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index% <> "")
            {
                tempExt:=ExtArr%A_Index%
                ErrorCount := CopyFiles(tempExt,sBackupPath,bRecursive)
                logErrors(tempExt, sBackupPath, ErrorCount)
                If (ErrorCount > 0)
                {
                    Return
                }
            }
        }
    } else If ( bCopyallExts = True)
     {
        ErrorCount := CopyFiles("*",sBackupPath,bRecursive)
        logErrors("*", sBackupPath, ErrorCount)
        If (ErrorCount > 0)
        {
            Return
        }
    }       
    If (bZipBackup = 1) {
        zipBackup(sDest "\Backup_00")
    }
    Gosub, ToggleBackup
    return
}
    
GuiClose:
{
    Gosub, ExitSub
    Return
}

ToggleBackup:
{
    toggle := !toggle
    if (toggle) 
    {
        if(tInterval < 60000)
        {
            tInterval:=60000
        }
        SetTimer, Backup, %tInterval%
    }else  {
        logEditAdd("Auto backup stopped.")
        FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
        sMainLogPath := sDest . sMainLogName
        if FileExist(sMainLogPath)
            FileAppend ,`n%sNow% Auto backup stopped., %sMainLogPath%
        else
            FileAppend ,%sNow% Auto backup stopped., %sMainLogPath%
        SetTImer, Backup, Off
    }
    return
}

backup:
{
    sBackupPath := sDest . "\Backup_" . iBkupNum
    sBackupLogPath := sBackupPath . "\stbackup_log.txt"
    FileDelete, %sBackupLogPath%
    FileRemoveDir, %sBackupPath%, 1
    if(bCopyallExts = false)
    {
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index% <> "")
            {
                tempExt:=ExtArr%A_Index%
                ErrorCount := CopyFiles(tempExt,sBackupPath,bRecursive)
                logErrors(tempExt, sBackupPath, ErrorCount)
                If (ErrorCount > 0)
                {
                    Return
                }
            }
        }
    } else If ( bCopyallExts = True)
     {
        ErrorCount := CopyFiles("*",sBackupPath,bRecursive)
        logErrors("*", sBackupPath, ErrorCount)
        If (ErrorCount > 0)
        {
            Return
        }
    }  
    If (bZipBackup = 1) {
        zipBackup(sDest "\Backup_" iBkupNum)
    }
    iBkupNum := iBkupNum + 1
    if (iBkupNum > iBackupCount )
    {
        iBkupNum := 1
    }
    IniWrite, %iBkupNum%, STB_settings.ini, History, Next Backup Number
    Return
}

BKbtn:
{
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, Extstring ,, extsediVar,
    sPath := trimPath(sPath)
    sDest := trimPath(sDest)
    trimExts(Extstring)
    sExts := Extstring
    GuiControl,,extsediVar, %sExts%  
    stringSplit, ExtArr, Extstring ,`;,
    sPVar :=InStr(FileExist(sPath),"D")
    If (sPVar=0)
    {
        CtlColors.Change(HSLedit, "FFC0C0", "")
        GuiControl,Focus, SLedit
        return
    }
    sCVar :=InStr(FileExist(sCustomDest),"D")
    If (sCVar!=0)
        FileSelectFolder,OutputVar3 ,*%sCustomDest% , 3, Manual backup location
    Else
        FileSelectFolder,OutputVar3 ,*%sDest% , 3, Manual backup location
    if OutputVar3 =
        return
    if (OutputVar3 = sPath)
    {
        OutputVar3 .= "\ST_Backups"
    }
    sDVar :=InStr(FileExist(OutputVar3),"D")
    If (sDVar=0)
    {
        FileCreateDir, %OutputVar3%
        erl:=ErrorLevel
        if(erl<>0)
        {
            SplashTextOff
            msgbox,% errIcon,, The backup path could not be created: %OutputVar3%
            return
        }
    }
    GuiControl,, BLedit, %sDest%
    GuiControl,, SLedit, %sPath%
    bCopyallExts:=false
    loop, %ExtArr0%
    {
        if(ExtArr%A_Index%="*")
        {
            bCopyallExts:=true
            Break
        }
    }
    sMainLogPath := sDest . sMainLogName
    FormatTime, sNow, %a_now% T12, [yyyy-MM-dd_HH-mm-ss]
    SplitPath, sPath, dname
    sBackupPath := OutputVar3 . "\STBackup_" . dname . "_" . sNow
    sBackupLogPath := sBackupPath . "\stbackup_log.txt"
    FileDelete, %sBackupLogPath%
    FileRemoveDir, %sBackupPath%, 1
    if(bCopyallExts = false)
    {
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index% <> "")
            {
                tempExt:=ExtArr%A_Index%
                ErrorCount := CopyFiles(tempExt,sBackupPath,bRecursive)
                logErrors(tempExt, sBackupPath, ErrorCount, false)
                If (ErrorCount > 0)
                {
                    Return
                }
            }
        }
    } else If ( bCopyallExts = True)
     {
        ErrorCount := CopyFiles("*",sBackupPath,bRecursive)
        logErrors("*", sBackupPath, ErrorCount, false)
        If (ErrorCount > 0)
        {
            Return
        }
    }
    If (bZipBackup = 1) {
        zipBackup(sBackupPath)
    }
    sCustomDest := OutputVar3
    IniWrite, %sCustomDest%, STB_settings.ini, History, Last Manual Backup Location
    SplashTextOff
    msgBox ,% infoIcon,, Backup finished.
    return
}

ExitSub:
{
    if A_ExitReason not in Logoff,Shutdown
    {
        SetTImer, Backup, Off
        sleep, 50
        GuiControlGet, Extstring ,, extsediVar,
        trimExts(Extstring)
        if(Extstring ="")
        {
            sExts := "*;"
        }else  {
            sExts := Extstring
        }
        if (_sPath<>"" and (StrReplaceVar(_sPath) = sPath)) {
            IniWrite, %_sPath%, STB_settings.ini, Paths, Files Location
        }
        else {
            IniWrite, %sPath%, STB_settings.ini, Paths, Files Location            
        }
        if (_sDest<>"" and ((StrReplaceVar(_sDest) = sDest))) {
            IniWrite, %_sDest%, STB_settings.ini, Paths, Backups Location
        }
        else {
            IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location            
        }
        IniWrite, %tInterval%, STB_settings.ini, Option, Backup Interval 
        IniWrite, %iBackupCount%, STB_settings.ini, Option, Backups Count 
        IniWrite, %iBkupNum%, STB_settings.ini, History, Next Backup Number
        IniWrite, %sExts%, STB_settings.ini, Option , Extensions
        IniWrite, %sCustomDest%, STB_settings.ini, History, Last Manual Backup Location
        IniWrite, %bZipBackup%, STB_settings.ini, Option, Zip Backups
        IniWrite, %bRecursive%, STB_settings.ini, Option, Recursive
        IniWrite, %bLogVarEnabled%, STB_settings.ini, Option, Show log
        IniWrite, %iMaxLogSize%, STB_settings.ini, Option, Max Log Size
        FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
        FileAppend ,`n%sNow% exiting program..., %sLogfullpath%
        sleep, 50
    }
    ExitApp
}