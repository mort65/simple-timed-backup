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
sCustomDest := ""
sExts := ""
iBackupCount := 10
tInterval := 300000 ; 5 min
toggle := 0
sCurrentTime :=""
bCopyallExts:=false
bRecursive:=false
red:="c0xe1256b"
blue:="c0x056bed"
bZipBackup := 0
errIcon := 16
infoIcon := 64
curVersion:=1.1
myName:="Simple Timed Backup"
myTitle:=myName . " " . curVersion
bLogVarEnabled := False

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
    Zip(sPath , sParent "\" sName ".zip")
    FileRemoveDir, %sPath%, 1
    FileCreateDir, %sPath%
    FileMove, %sParent%\%sName%.zip,%sPath%,1
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
}

logErrors(sExt,sBackupPath,errCount,bSilent:=true)
{
    global sPath
    global sDest
    global mainStatusBarVar_TT
    sMainLogPath := sDest
    sMainLogPath .= "\stbackup_log.txt"
    sBackupLogPath := sBackupPath
    sBackupLogPath .= "\stbackup_log.txt"
    FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
    FormatTime, curTime, %a_now% T12, [HH:mm]
    if (errCount < 0)
    {
        ;No file from Source folder copied without any error.
        return
    }
    else if (errCount = 0)
    {
        FormatTime, sCurrentTime ,  dddd MMMM d yyyy HH:mm:ss T12
        if FileExist(sMainLogPath)
        {
            FileGetSize, logsizekb, %sMainLogPath%, K
            if(logsizekb>500)
            {
                FileDelete, %sMainLogPath%
                FileAppend ,%sNow% backup started..., %sMainLogPath%
                FileAppend ,`n%sNow% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            }else  {
                FileAppend ,`n%sNow% backup started..., %sMainLogPath%
                FileAppend ,`n%sNow% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            }
        }else  {
            FileAppend ,%sNow% backup started..., %sMainLogPath%
            FileAppend ,`n%sNow% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
        }
        FileDelete, %sBackupLogPath%       
        FileAppend ,*.%sExt% Backup: in %sCurrentTime%,%sBackupLogPath%
        mainStatusBarVar_TT := sBackupPath
        strLog := "Backup: """ . trimPath(shrinkString(sBackupPath,62,"m")) . """"
        SB_SetText(A_Tab  . curTime . " " . strLog,1,1)
        logEditAdd(strLog)
    } else {
        if FileExist(sMainLogPath)
        {
            FileAppend ,`n%sNow% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            FileAppend ,`n%sNow% can`t copy %errCount% file(s)!
        }else  {
            FileAppend ,%sNow% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            FileAppend ,`n%sNow% can`t copy %errCount% file(s)!
        }
        strLog := "Error: Cannot copy " . errCount " files!"
        ;SB_SetText(A_Tab  . curTime . " " . strLog,1,1)
        logEditAdd(strLog)
        if (bSilent=true)
        {
            return
        }
        SplashTextOff
        msgBox,% errIcon,, Cannot copy some files!
        Gosub, ExitSub
        return
    }
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
    trimPath(sPath)
    trimPath(sDest)
    trimPath(sDestFolder)
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
    IniRead, iBkupNum, STB_settings.ini, History, Next Backup Number, 1
    IniRead, sExts, STB_settings.ini, Option , Extensions, "*;"
    IniRead, bZipBackup, STB_settings.ini, Option, Zip Backups , 0
    IniRead, bRecursive, STB_settings.ini, Option, Recursive , 0
    IniRead, bLogVarEnabled, STB_settings.ini, Option, Show log , 0
    sPath := StrReplaceVar(trimPath(_sPath))
    sDest := StrReplaceVar(trimPath(_sDest))
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
Gui,Font, normal s8, %_font%
Gui,Add,Text,x5 y12 w85 %black% left, Files to Backup
Gui,Add,Text,x5 y47 w85 %black% left, Backups Location
Gui,Add,Edit,x95 y10 w332 h30 HwndHSLedit r1 %black%  ReadOnly  vSLedit gRevertSLeditColor,
GuiControl,, SLedit, %sPath%
Gui,Add,Edit,x95 y45 w332 h30 HwndHBLedit r1 %black% ReadOnly  vBLedit gRevertBLeditColor,
Gui,Font, s8 normal, %_font%
GuiControl,, BLedit, %sDest%
Gui,Add,Button,x437 y9 w30 h23 r1 center vSPvar gSPbtn,...
Gui,Add,Button,x437 y44 w30 h23 r1 center vBPvar gBPbtn,...
Gui,Add, GroupBox, x5 y80 w310 h101, Backup Options
Gui,Add,Text,x10 y97 w80 h13 left  %black%  ,Backup every
Gui,Add,Edit,x85 y95 w70 h18 %black% Number ReadOnly  vBIedit gBIedit
mInterval := (tInterval/60000)
Gui,Add,UpDown, 0x20  Range1-720 ,%mInterval%,vBIud
Gui,Add,Text,x10 y116 w80 h13 %black% left ,Backup count

Gui,Add,Edit,x85 y114 w70 h18 %black% Number ReadOnly  vBCedit gBCedit
Gui,Add,UpDown, 0x20  Range1-100,%iBackupCount%,vBCud

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

Gui,Add,Button, x165 y139 w147 h35 center +Disabled  vBKvar gBKbtn , Manual Backup

Gui,Add,Button,x320 y97 w147 h35 +Disabled vDEvar gDEbtn,Deactivate
Gui,Add,Button,x320 y139 w147 h35 center vACvar gACbtn,Activate
Gui,Font, s8 normal, %_font%
Gui,Add, Checkbox, x5 y182 vShowLogcbx gToggleLogcbx, Show Log

if sPath !=
    GuiControl, Enabled, BKvar

Gui,Font, s8 normal, %_font%
Gui, Add, StatusBar,gmainStatusBar vmainStatusBarVar,Ready
Gui,Font, s7 , Lucida Console
Gui,add, edit, x9 y202 w468 r9 left ReadOnly vLogEditVar gLogEdit
Gui,Font, s8 normal, %_font%

if (bLogVarEnabled=False)
{
    GuiControl,, ShowLogcbx, 0
    GuiControl, Hide, LogEditVar
    Gui,Show, h224,%myTitle%
}
else
{
    GuiControl,, ShowLogcbx, 1
    GuiControl, Show, LogEditVar
    Gui,Show, h334,%myTitle%
  
}

;Tooltips
SLedit_TT := "The source folder."
BLedit_TT := "The destination folder for storing backups."
SPvar_TT := "Change the source folder."
BPvar_TT := "Change the destination folder."
BCedit_TT := "How many backups should be created before overwriting previous backups."
ACvar_TT := "First, a backup will be created inside the ""backup_0"" folder.`nThen automated backups will be created at the selected interval."
DEvar_TT := "First, a backup will be created inside the ""backup_0"" folder.`nThen creating automated backups will be stopped."
extsediVar_TT := "Extensions are separated by `;`n* means any extension"
BKvar_TT := "Takes a manual backup inside the selected folder."
ZipBackupvar_TT := "Toggles the compression of backups."
RecursiveVar_TT := "Toggles backup for files in subfolders."
BIedit_TT := "Automated backups will be created after the selected minutes."
EDbtnvar_TT := "Edit what file types to backup."
mainStatusBarVar_TT := ""
LogEditVar_TT := ""

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
    global mainStatusBarVar_TT
    if (mainStatusBarVar_TT<>"")
    {
        if InStr(FileExist(mainStatusBarVar_TT),"D")
        {
            Run, Explorer /n`,/e`,%mainStatusBarVar_TT%
        }
    }
    return
}

ToggleLogcbx:
{
    Global bLogVarEnabled
    GuiControlGet, bLogVarEnabled, , ShowLogcbx
    If (bLogVarEnabled=False)
    {
      GuiControl, Hide, LogEditVar
      Gui,Show, h224
    }
    Else
    {
      GuiControl, Show, LogEditVar
      Gui,Show, h334
      
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
    IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location
    Return
}

BIedit:
{
    GuiControlGet , BIedit
    tInterval := BIedit*60000
    IniWrite, %tInterval%, STB_settings.ini, Option, Backup Interval 
    Return
}

BCedit:
{
    GuiControlGet , BCedit
    iBackupCount := BCedit
    IniWrite, %iBackupCount%, STB_settings.ini, Option, Backups Count
    Return
}
    
BCud:
{
    GuiControlGet , BCud
    iBackupCount := BCud
    IniWrite, %iBackupCount%, STB_settings.ini, Option, Backups Count
    Return
}

BIud:
{
    GuiControlGet , BIud
    tInterval := BIud*60000
    IniWrite, %tInterval%, STB_settings.ini, Option, Backups Interval 
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
    {
        GuiControl,,extsediVar, %sExts%
    }   
    Gui, Submit , NoHide
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, tInterVal,, BIedit
    GuiControlGet, iBackupCount,, BCedit
    GuiControlGet, Extstring ,, extsediVar,
    trimExts(Extstring)
    sExts := Extstring
    StringSplit, ExtArr, Extstring ,`;,
    PathPattern := spath
    sPVar :=InStr(FileExist(PathPattern),"D")
    if(tInterval="" )
    {
        tInterval := 300000
        GuiControl, , BIud,%tInterval%
    }else   If (iBackupCount="")
     {
        iBackupCount := 10
    }else if (sPVar=0)
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
    }else  
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
        sBackupPath = %sDest%\Backup_0
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
    logEditAdd("Auto backup stopped.")
    sBackupPath = %sDest%\Backup_0
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
        SetTImer, Backup, Off
    }
    return
}

backup:
{
    sBackupPath = %sDest%\Backup_%iBkupNum%
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
    sCVar :=InStr(FileExist(sCustomDest),"D")
    If (sCVar!=0)
        FileSelectFolder,OutputVar3 ,*%sCustomDest% , 3, Manual backup location
    Else
        FileSelectFolder,OutputVar3 ,*%sDest% , 3, Manual backup location
    if OutputVar3 =
    return
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, Extstring ,, extsediVar, 
    trimExts(Extstring)
    sExts := Extstring
    stringSplit, ExtArr, Extstring ,`;,
    sPVar :=InStr(FileExist(sPath),"D")
    If (sPVar=0)
    {
        SplashTextOff
        msgbox,% errIcon,, The path could not be found: %sPath%
        return
    }
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
    bCopyallExts:=false
    loop, %ExtArr0%
    {
        if(ExtArr%A_Index%="*")
        {
            bCopyallExts:=true
            Break
        }
    }
    FormatTime, sNow, %a_now% T12, [yyyy-MM-dd_HH-mm-ss]
    SplitPath, sPath, dname
    sBackupPath = %OutputVar3%\STBackup_%dname%_%sNow%
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
        FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
        FileAppend ,`n%sNow% exiting program..., %sLogfullpath%
        sleep, 50
    }
    ExitApp
}