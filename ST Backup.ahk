#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent 
#SingleInstance, ignore 
#NoTrayIcon

#Include Class_Util.ahk
#Include Class_CtlColors.ahk
#Include Class_Ini.ahk

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
bCopyallExts := false
bRecursive := false
errIcon := 16
infoIcon := 64
curVersion := 1.117
myName :="Simple Timed Backup"
_WinH := 328
_WinW := 635
iMaxLogSize := 500 ;kb
_backup_ext := ".stb.zip"
bInfiniteBkup := false
iBkupNum := 1
_bExiting := False
iniPath := A_ScriptDir . "\STB_settings.ini"

_font :="Tahoma"

myTitle := myName " " util.getVersion(curVersion)

AutoSize(DimSize, cList*) {
    Static cInfo := {}
    Local

    If (DimSize = "reset") {
        Return cInfo := {}
    }

    For i, ctrl in cList {
        ctrlID := A_Gui . ":" . ctrl
        If (cInfo[ctrlID].x = "") {
            GuiControlGet i, %A_Gui%: Pos, %ctrl%
            MMD := InStr(DimSize, "*") ? "MoveDraw" : "Move"
            fx := fy := fw := fh := 0
            For i, dim in (a := StrSplit(RegExReplace(DimSize, "i)[^xywh]"))) {
                If (!RegExMatch(DimSize, "i)" . dim . "\s*\K[\d.-]+", f%dim%)) {
                    f%dim% := 1
                }
            }

            If (InStr(DimSize, "t")) {
                GuiControlGet hWnd, %A_Gui%: hWnd, %ctrl%
                hWndParent := DllCall("GetParent", "Ptr", hWnd, "Ptr")
                VarSetCapacity(RECT, 16, 0)
                DllCall("GetWindowRect", "Ptr", hWndParent, "Ptr", &RECT)
                DllCall("MapWindowPoints", "Ptr", 0, "Ptr"
                , DllCall("GetParent", "Ptr", hWndParent, "Ptr"), "Ptr", &RECT, "UInt", 1)
                ix -= (NumGet(RECT, 0, "Int") * 96) // A_ScreenDPI
                iy -= (NumGet(RECT, 4, "Int") * 96) // A_ScreenDPI
            }

            cInfo[ctrlID] := {x: ix, fx: fx, y: iy, fy: fy, w: iw, fw: fw, h: ih, fh: fh, gw: A_GuiWidth, gh: A_GuiHeight, a: a, m: MMD}

        } Else If (cInfo[ctrlID].a.1) {
            dgx := dgw := A_GuiWidth - cInfo[ctrlID].gw
            dgy := dgh := A_GuiHeight - cInfo[ctrlID].gh

            Options := ""
            For i, dim in cInfo[ctrlID]["a"] {
                Options .= dim . (dg%dim% * cInfo[ctrlID]["f" . dim] + cInfo[ctrlID][dim]) . A_Space
            }

            GuiControl, % A_Gui ":" cInfo[ctrlID].m, % ctrl, % Options
        }
    }
}


zipBackup(sPath)
{
    if InStr(FileExist(sPath), "D") {
        if util.IsEmpty(sPath) {
            return
        }
    }
    else {
        return
    }
    SplitPath, sPath, sName, sParent
    FileDelete, %sParent%\%sName%.stb.zip
    FileDelete, %sPath%\%sName%.stb.zip
    sOut := sParent "\" sName ".stb.zip"
    psOut := util.psEscape(sOut)
    psPath := util.psEscape(sPath)
    Run PowerShell.exe -Command (Compress-Archive -Path '%psPath%\*' -CompressionLevel Optimal -DestinationPath '%sOut%'); if ($?) {(Remove-Item -force '%psPath%' -recurse -Confirm:$False);},, Hide UseErrorLevel
    if (ErrorLevel = "ERROR")
    {
        util.Zip(sPath , sParent "\" sName ".stb.zip")      
        FileRemoveDir, %sPath%, 1
    }
    Return
}

trimExts(ByRef sExtensions)
{
    sExtensions:=StrReplace(sExtensions, "`r`n")
    sExtensions:=StrReplace(sExtensions, "`n")
    sExtensions:=StrReplace(sExtensions, A_SPACE)
    sExtensions:=StrReplace(sExtensions, A_Tab)
    sExtensions:=StrReplace(sExtensions, ".")
    sExtensions:=StrReplace(sExtensions, "/")
    sExtensions:=StrReplace(sExtensions, "\")
    sExtensions:=StrReplace(sExtensions, ":")
    sExtensions:=StrReplace(sExtensions, "|")
    sExtensions:=StrReplace(sExtensions, """")
    sExtensions:=StrReplace(sExtensions, "<")
    sExtensions:=StrReplace(sExtensions, ">")
    sExtensions:=StrReplace(sExtensions, ",")
    sExtensions:=StrReplace(sExtensions, "?")    
    if(sExtensions ="")
    {
        sExtensions := "*;"
    }
}

logErrors(sExt,sBackupPath,errCount,bSilent:=true)
{
    global sPath
    global sDest
    Global sBackupLogPath
    Global sMainLogName
    Global iMaxLogSize
    Global sMainLogPath
    Global _backup_ext
    Global errIcon
    FormatTime, sTime, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
    if (errCount < 0)
    {
        sLog := "Warning: No file copied. Type=*." . sExt
        logEditAdd(sLog)
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
                FileAppend ,%sTime% backup started..., %sMainLogPath%
                FileAppend ,`n%sTime% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%%_backup_ext%, %sMainLogPath%
            } 
            else  
            {
                FileAppend ,`n%sTime% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%%_backup_ext%, %sMainLogPath%
            }
        } 
        else  
        {
            FileAppend ,%sTime% backup started..., %sMainLogPath%
            FileAppend ,`n%sTime% backup: `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%%_backup_ext%, %sMainLogPath%
        }           
        sLog := "*." . sExt . " Backup: """ . util.trimPath(sBackupPath . _backup_ext) . """"
        logEditAdd(sLog)
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
            FileAppend ,`n%sTime% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%%_backup_ext%, %sMainLogPath%
            FileAppend ,`n%sTime% can`t copy %errCount% file(s)!
        }
        else
        {
            FileAppend ,%sTime% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%%_backup_ext%, %sMainLogPath%
            FileAppend ,`n%sTime% can`t copy %errCount% file(s)!
        }
        sLog := "Error: Cannot copy " . errCount . " file(s) to destination. Type=*." . sExt
        logEditAdd(sLog)
        if (!bSilent) 
        {            
            msgBox,% errIcon,, Cannot copy some files!
            ExitFunc()
        }
    }
    Return
}

CopyFiles(sExt,sDestFolder,sSourceFolder:="",bRecursive:=false)   
{
    global sPath
    global sDest
    sPath:= util.trimPath(sPath)
    sDest:= util.trimPath(sDest)
    sDestFolder:= util.trimPath(sDestFolder)
    if !sSourceFolder
        sSourceFolder:=sPath
    if util.IsEmpty(sSourceFolder)
    {
        return -1
    }
    FileCreateDir, %sDestFolder%
    ErrCount:=ErrorLevel
    if(ErrCount<>0)
    {
        return
    }
    ErrCount := 0
    if (bRecursive=false) {
        FileCopy, %sSourceFolder%\*.%sExt%, %sDestFolder%\, 1
        ErrCount := ErrorLevel
        If (ErrCount > 0)
        {
            Return ErrCount
        }
    } else {
        Loop Files, %sSourceFolder%\*.%sExt%, R  ; Recurse into subfolders.
        {
            if (A_LoopFileDir=sSourceFolder)
            {
                FileCopy, %A_LoopFileFullPath%, %sDestFolder%\, 1
            }
            else
            {
                sDestFileLongPath := StrReplace(A_LoopFileLongPath,sSourceFolder,sDestFolder)
                SplitPath, sDestFileLongPath,, sDestFileDir
                If (util.bIsParentPath(sDest,A_LoopFileLongPath))
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
            ErrCount := ErrorLevel
            If (ErrCount > 0)
            {
                Return ErrCount
            }
        }
    }
    if util.IsEmpty(sDestFolder)
    {
        return -1
    }
    return 0
}

if (FileExist("STB_settings.ini"))
{
    sleep, 50
    myINI := new Ini(iniPath)
    util.setVar(_sPath,Paths_FilesLocation)
    util.setVar(_sDest,Paths_BackupsLocation)
    util.setVar(sCustomDest,History_LastManualBackupLocation)
    util.setVar(tInterval,Option_BackupInterval,300000)
    util.setVar(iBackupCount,Option_BackupsCount,10)
    util.setVar(iMaxLogSize,Option_MaxLogSize,500)
    util.setVar(iBkupNum,History_NextBackupNumber,1)
    util.setVar(sExts,Option_Extensions,"*;")
    util.setVar(bRecursive,Option_Recursive)
    util.setVar(bInfiniteBkup,Option_UnlimitedBackups)
    sPath := util.StrReplaceVar(util.trimPath(_sPath))
    sDest := util.StrReplaceVar(util.trimPath(_sDest))
    if (sPath<>"") 
    {
        if (sPath=sDest) 
        {
            sDest .= "\ST_Backups"
            myINI.iniEdit("Paths","Backups Location",sDest)
            
        }        
    }
    sMainLogPath := sDest . sMainLogName
    sCustomDest := util.trimPath(sCustomDest)
    if (sCustomDest<>"" and sCustomDest=sPath) 
    {
        sCustomDest .="\ST_Backups"
        myINI.iniEdit("History","Last Manual Backup Location",sCustomDest)
    }
}
else 
{
    sExts:= "*;"
    str := "[Paths]"
    str := str . "`nFiles Location=" . sPath
    str := str . "`nBackups Location=" . sDest
    str := str . "`n[History]"
    str := str . "`nLast Manual Backup Location=" . sCustomDest
    str := str . "`nNext Backup Number=" . iBkupNum
    str := str . "`n[Option]"
    str := str . "`nBackup Interval=" . tInterval
    str := str . "`nBackups Count=" . iBackupCount
    str := str . "`nMax Log Size=" . iMaxLogSize
    str := str . "`nExtensions=" . sExts
    str := str . "`nRecursive=" . bRecursive
    str := str . "`nUnlimited Backups=" . bInfiniteBkup
    str := Trim(str,"`n`r `t")
    FileAppend, %str%,%iniPath%,UTF-16 
    sleep, 50
    myINI := new Ini(iniPath)
}

Hotkey, ^!x, ExitFunc
OnExit("ExitFunc")

Gui +LastFound
Gui, -ToolWindow
Gui, +CAPTION
Gui, -MaximizeBox
Gui, +Resize +MinSize
Gui, Margin,11,
Gui,Font, normal s8, %_font%
Gui,Add,Text,x9 y12 w85 left, Files to Backup
Gui,Add,Text,x9 y47 w85 left, Backups Location
Gui,Add,Edit,x95 y10 w487 h30 HwndHSLedit r1  vSLedit gRevertSLeditColor,
GuiControl,, SLedit, %sPath%
Gui,Add,Edit,x95 y45 w487 h30 HwndHBLedit r1  vBLedit gRevertBLeditColor,
Gui,Font, s8 normal, %_font%
GuiControl,, BLedit, %sDest%
Gui,Add,Button,x592 y9 w30 h23 r1 center vSPvar gSPbtn,...
Gui,Add,Button,x592 y44 w30 h23 r1 center vBPvar gBPbtn,...
Gui,Add, GroupBox, x5 y80 w465 h101 vBSGbx,
Gui,Add,Text,x10 y97 w80 h13 left ,Backup every
Gui,Add,Edit,x85 y95 w70 h18  number vBIedit gBIedit
mInterval := (tInterval//60000)
Gui,Add,UpDown, 0x20  Range1-720 ,%mInterval%,vBIud
Gui,Add,Text,x10 y116 w80 h13  left ,Backup count

Gui,Add,Edit,x85 y114 w70 h18 Number vBCedit gBCedit
Gui,Add,UpDown, 0x20  Range1-10000,%iBackupCount%,vBCud
Gui,Add, GroupBox, x10 y134 w306 h41 vBFGbx, Backup these file types
Gui,Add,Edit,x15 y150 w296 h20  Lowercase vextsediVar gextsEdit,%sExts%
Gui,Add,Checkbox,x166 y115 w70 h20 -Wrap  vInfiniteBkupVar gInfiniteBkupcbx, Unlimited

if (bInfiniteBkup)
{
    GuiControl,, InfiniteBkupVar, 1
    GuiControl,Disable,BCedit
}
else
{
    GuiControl,, InfiniteBkupVar, 0
    GuiControl,Enable,BCedit
}

Gui,Add,Checkbox,x242 y115 w70 h20 -Wrap  vRecursiveVar gRecursivecbx,Recursive

if (bRecursive)
{
    GuiControl,, RecursiveVar, 1    
}
else
{
    GuiControl,, RecursiveVar, 0
}

Gui,Add,Button, x320 y92 w147 h35 center +Disabled  vBKvar gBKbtn , Back up...

Gui,Add,Button, x320 y139 w147 h35 center +Disabled  vRSvar gRSbtn , Restore...

Gui,Add,Button,x475 y92 w147 h35 +Disabled vDEvar gDEbtn,Deactivate
Gui,Add,Button,x475 y139 w147 h35 center vACvar gACbtn,Activate
Gui,Font, s8 normal, %_font%
Gui,Add,Text,x166 y97 w80 h13 left ,Max Log Size
Gui,Add,Edit,x242 y95 w70 h18 number vLSedit gLSedit
Gui,Add,UpDown, 0x20 Range10-100000 vLSud,%iMaxLogSize%

if sPath !=
    {
        GuiControl, Enabled, RSvar
        if sDest !=
        {
            GuiControl, Enabled, BKvar  
        }
    }

Gui,Font, s8 normal, %_font%
Gui, Add, StatusBar,gmainStatusBar vmainStatusBarVar,%A_Tab%Ready
Gui,Font, s8 , Lucida Console
Gui,add, edit, x9 y203 w614 h98 r9 left ReadOnly vLogEditVar gLogEdit
Gui,Font, s8 normal, %_font%
;Gui,Show, w%_WinW% h%_WinH% center ,%myTitle%
Gui,Show, autoSize center ,%myTitle%


;Tooltips
SLedit_TT := "The source folder. Double-click to open."
BLedit_TT := "The destination folder for storing backups. Double-click to open."
SPvar_TT := "Change the source folder."
BPvar_TT := "Change the destination folder."
BCedit_TT := "How many backups should be created before overwriting previous backups."
ACvar_TT := "First, a backup will be created inside the ""backup_0"" folder.`nThen automated backups will be created at the selected interval."
DEvar_TT := "First, a backup will be created inside the ""backup_00"" folder.`nThen creating automated backups will be stopped."
extsediVar_TT := "Extensions are separated by `;`n* means any extension"
BKvar_TT := "Takes a manual backup inside the selected folder."
RSvar_TT := "Restore a backup archive to the source folder."
ZipBackupvar_TT := "Toggles the compression of backups."
RecursiveVar_TT := "Toggles backup for files in subfolders."
InfiniteBkupVar_TT := "Toggles unlimited backups without overwriting previous ones."
BIedit_TT := "The time between auto-updates in minutes."
EDbtnvar_TT := "Edit what file types to backup."
mainStatusBarVar_TT := ""
LogEditVar_TT := "Double-click to open the log file if exists."
LSedit_TT := "Max allowed size of the log file in Kbytes."

OnMessage(0x200, "WM_MOUSEMOVE")
OnMessage(0x0203, "WM_LBUTTONDBLCLK")
Return

Resize:
{
	AutoSize("reset") ; Needs to reset if you changed the Control size manually.
    return
}

GuiSize:
{
    If (A_EventInfo = 1) ; The window has been minimized.
		Return
    AutoSize("x", "SPvar")
    AutoSize("w", "SLedit")
    AutoSize("x", "BPvar")
    AutoSize("w", "BLedit")
    AutoSize("w", "BSGbx")
    AutoSize("w", "BSGbx")
    AutoSize("w", "BFGbx")
    AutoSize("x", "BKvar")
    AutoSize("x", "RSvar")
    AutoSize("x", "DEvar")
    AutoSize("x", "ACvar")
    AutoSize("w h0.99", "LogEditVar")
    AutoSize("w","extsediVar")
    Return
}

WM_MOUSEMOVE()
{
    static PrevControl := ""
    static _TT := "" ; _TT is kept blank for use by the ToolTip command below.
    static CurrControl := ""
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
            if (!FileExist(logPath))
            {
                FileAppend,, %logPath%
            }
            SplitPath, logPath,,,,fName 
            Run,% "notepad.exe " . logPath
            If WinExist(fName)
            {
                WinActivate
            }
        }

    }
    return
}


resetGUI:
{
    Gui -Disabled
    Gui,Show, autoSize ,%myTitle%
    SB_SetText(A_Tab  . "ready",1,1)
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
    Gui +Disabled
    FileSelectFolder,OutputVar1 ,*%sPath% , 0, Files location
    Gosub, resetGUI   
    
    if OutputVar1 =
        return
    GuiControl,, SLedit, %OutputVar1%
    GuiControl, Enabled, BKvar
    GuiControl, Enabled, RSvar
    sPath := OutputVar1
    myINI.iniEdit("Paths","Files Location",sPath)
    Return
}

BPbtn:
{
    Gui +Disabled
    FileSelectFolder,OutputVar2 ,*%sDest% , 3, Backups location
    Gosub, resetGUI
    
    if OutputVar2 =
        return
    if (OutputVar2 = sDest)
        return
    FileCreateDir, %OutputVar2%\ST_Backups
    OutputVar2 .= "\ST_Backups"
    GuiControl,, BLedit, %OutputVar2%
    sDest := OutputVar2
    sMainLogPath := sDest . sMainLogName
    GuiControl, Enabled, BKvar
    myINI.iniEdit("Paths","Backups Location",sDest)
    Return
}

BIedit:
{
    GuiControlGet ,InVar,, BIedit
    util.checkNum(InVar, "BIedit", 1,720)
    tInterval := InVar*60000
    myINI.iniEdit("Option","Backup Interval",tInterval)
    Return
}

BCedit:
{
    GuiControlGet ,InVar,, BCedit
    util.checkNum(InVar, "BCedit", 1,10000)
    iBackupCount := InVar
    myINI.iniEdit("Option","Backups Count",iBackupCount)
    Return
}

LSedit:
{
    GuiControlGet, InVar,, LSedit
    util.checkNum(InVar, "LSedit", 10,100000)
    iMaxLogSize := InVar
    myINI.iniEdit("Option","Max Log Size",iMaxLogSize)
    Return
}

extsEdit:
{
    Return
}
    
Recursivecbx:
{
    bRecursive := !bRecursive
    myINI.iniEdit("Option","Recursive",bRecursive)
    Return    
}


InfiniteBkupcbx:
{
    bInfiniteBkup := !bInfiniteBkup
    if (bInfiniteBkup)
    {
        GuiControl,Disable,BCedit
    }
    Else
    {
        GuiControl,Enable,BCedit
    }
    myINI.iniEdit("Option","Unlimited Backups",bInfiniteBkup)
    Return 
}

    
ACbtn:
{ 
    Gui, Submit , NoHide
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, mInterval,, BIedit
    GuiControlGet, iBackupCount,, BCedit
    GuiControlGet, Extstring ,, extsediVar,
    trimExts(Extstring)
    sPath := util.trimPath(sPath)
    sDest := util.trimPath(sDest)
    sExts := Extstring
    GuiControl,,extsediVar, %sExts% 
    StringSplit, ExtArr, Extstring ,`;,
    sPVar :=InStr(FileExist(spath),"D")
    if(mInterval="" )
    {
        tInterval := 300000
        mInterval := 5
        GuiControl, , BIud,%mInterval%
    }
    else if (iBackupCount="")
    {
        iBackupCount := 10
        GuiControl, , BCud,%iBackupCount%
    }
    else if (sPVar=0)
    {
        CtlColors.Change(HSLedit, "FFC0C0", "")
        GuiControl,Focus, SLedit
        Return
    }
    Else if mInterval not between 1 and 720
    {        
        msgbox,% errIcon,, Backup Interval is not within the valid range: 1-720
        return
    }
    Else if (!bInfiniteBkup && (iBackupCount not between 1 and 10000))
    {
        msgbox,% errIcon,, Backup Count is not within the valid range: 1-10000
        return
    }
    Else if iMaxLogSize not between 1 and 100000
    {
        msgbox,% errIcon,, Max Log Size is not within the valid range: 1-100000
        return
    }
    else  
    {
        if (sPath=sDest)
        {
            if (_sPath<>"" and (_sPath=_sDest) and (util.StrReplaceVar(_sPath)=sPath) and (util.StrReplaceVar(_sDest)=sDest)) 
            {
                _sDest.="\ST_Backups"
                sDest.="\ST_Backups"
                myINI.iniEdit("Paths","Backups Location",_sDest)
            }
            else {
                sDest.="\ST_Backups"
                myINI.iniEdit("Paths","Backups Location",sDest)
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
                msgbox,% errIcon,, The path you entered could not be created: %sDest%
                return
            }
        }
        GuiControl,, BLedit, %sDest%
        GuiControl,, SLedit, %sPath%
        tInterval := mInterval * 60000
        GuiControl,Disable,ACvar
        GuiControl,Enable,DEvar
        GuiControl,Disable,RSvar
        GuiControl,Disable,SPvar
        GuiControl,Disable,BPvar
        GuiControl,Disable,InfiniteBkupVar
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
        }else if (!bInfiniteBkup) {
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
        If (bCopyallExts)
         {
            ErrorCount := CopyFiles("*",sBackupPath,,bRecursive)
            logErrors("*", sBackupPath, ErrorCount)
            If (ErrorCount > 0)
            {
                Return
            }
        }
        else
        {
            loop, %ExtArr0%
            {
                if(ExtArr%A_Index% <> "")
                {
                    tempExt:=ExtArr%A_Index%
                    ErrorCount := CopyFiles(tempExt,sBackupPath,,bRecursive)
                    logErrors(tempExt, sBackupPath, ErrorCount)
                    If (ErrorCount > 0)
                    {
                        Return
                    }
                }
            }
        }
        zipBackup(sDest "\Backup_0")        
        Gosub, ToggleBackup
    }
    Return
}

DEbtn:
{
    GuiControl,Disable,DEvar
    GuiControl,Enable,ACvar
    GuiControl,Enable,RSvar
    GuiControl,Enable,SPvar
    GuiControl,Enable,BPvar
    GuiControl,Enable,InfiniteBkupVar
    if (bInfiniteBkup)
    {
        GuiControl,Disable,BCedit
    }
    Else
    {
        GuiControl,Enable,BCedit
    }
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
    If (bCopyallExts)
     {
        ErrorCount := CopyFiles("*",sBackupPath,,bRecursive)
        logErrors("*", sBackupPath, ErrorCount)
        If (ErrorCount > 0)
        {
            Return
        }
    }
    else
    {
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index% <> "")
            {
                tempExt:=ExtArr%A_Index%
                ErrorCount := CopyFiles(tempExt,sBackupPath,,bRecursive)
                logErrors(tempExt, sBackupPath, ErrorCount)
                If (ErrorCount > 0)
                {
                    Return
                }
            }
        }
    }
    zipBackup(sDest "\Backup_00")
    Gosub, ToggleBackup
    return
}
    
GuiClose:
{
    ExitFunc()
    Return
}

ToggleBackup:
{
    toggle := !toggle
    if (toggle) 
    {
        if(tInterval < 60000)
        {
            tInterval:= 60000
            mInterval := 5
        }
        SetTimer, Backup, %tInterval%
    }else  {
        logEditAdd("Auto backup stopped.")
        SB_SetText(A_Tab  . "Auto backup stopped.",1,1)
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
    if (bInfiniteBkup)
    {
        While (FileExist(sDest . "\Backup_" . iBkupNum . _backup_ext))
            iBkupNum += 1
    }
    sBackupPath := sDest . "\Backup_" . iBkupNum
    sBackupLogPath := sBackupPath . "\stbackup_log.txt"
    FileDelete, %sBackupLogPath%
    FileRemoveDir, %sBackupPath%, 1
    If (bCopyallExts)
     {
        ErrorCount := CopyFiles("*",sBackupPath,,bRecursive)
        logErrors("*", sBackupPath, ErrorCount)
        If (ErrorCount > 0)
        {
            Return
        }
    }
    else 
    {
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index% <> "")
            {
                tempExt:=ExtArr%A_Index%
                ErrorCount := CopyFiles(tempExt,sBackupPath,,bRecursive)
                logErrors(tempExt, sBackupPath, ErrorCount)
                If (ErrorCount > 0)
                {
                    Return
                }
            }
        }
    }
    zipBackup(sDest "\Backup_" iBkupNum)
    iBkupNum := iBkupNum + 1
    if (!bInfiniteBkup && (iBkupNum > iBackupCount ))
    {
        iBkupNum := 1
    }
    myINI.iniEdit("History","Next Backup Number",iBkupNum)
    Return
}

RSbtn:
{
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, Extstring ,, extsediVar,
    sPath := util.trimPath(sPath)
    sDest := util.trimPath(sDest)
    trimExts(Extstring)
    sExts := Extstring
    sPVar :=InStr(FileExist(sPath),"D")
    If (sPVar=0)
    {
        CtlColors.Change(HSLedit, "FFC0C0", "")
        GuiControl,Focus, SLedit
        return
    }
    sCVar :=InStr(FileExist(sCustomDest),"D")
    Gui +Disabled
    SelectedFile:=""
    If (sCVar!=0)
        FileSelectFile, SelectedFile, 3,%sCustomDest% , Open a file, Backup archive (*.stb.zip)
    Else
        FileSelectFile, SelectedFile, 3, %sDest%, Open a file, Backup archive (*.stb.zip)
    Gui -Disabled
    Gui,Show, autoSize center ,%myTitle%
    if (SelectedFile = "")
        Return
    shortFile := util.shrinkString(SelectedFile,55,"l")
    Gui +Disabled
    MsgBox 547,Restore, Are you sure you want to restore this backup?`n`r"%shortFile%"
    IfMsgBox No
    {
        Gosub, resetGUI
        Return
    } 
    IfMsgBox Cancel
    {
        Gosub, resetGUI
        Return
    }
        
    SB_SetText(A_Tab  . "Restoring...",1,1)
    tempPath:= sPath . "\._sb_restore"
    FileRemoveDir, %tempPath%, 1
    FileCreateDir, %tempPath%
    util.Unz(SelectedFile, tempPath)
    if not FileExist(tempPath . sMainLogName)
    {
        FileRemoveDir, %tempPath%, 1
        msgBox ,% errIcon,, Invalid Backup file!
        Gosub, resetGUI
        Return            
    }
    FileDelete, %tempPath%%sMainLogName%
    CopyFiles("*",sPath,tempPath,true)
    FileRemoveDir, %tempPath%, 1
    msgBox ,% infoIcon,, Restore finished.
    Gosub, resetGUI
    strLog := "Restore: """ . util.trimPath(SelectedFile) . """"
    logEditAdd(strLog)
    sMainLogPath := sDest . sMainLogName
    FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
    if (FileExist(sMainLogPath)) 
    {
        FileAppend ,`n%sNow% Restored: %SelectedFile%,%sMainLogPath%
    }
    else
    {
        FileAppend ,%sNow% Restored: %SelectedFile%,%sMainLogPath%
    } 
    ;Run, Explorer /n`,/e`,%sPath%
    return
}

BKbtn:
{
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, Extstring ,, extsediVar,
    sPath := util.trimPath(sPath)
    sDest := util.trimPath(sDest)
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
    Gui +Disabled
    If (sCVar!=0)
        FileSelectFolder,OutputVar3 ,*%sCustomDest% , 3, Manual backup location
    Else
        FileSelectFolder,OutputVar3 ,*%sDest% , 3, Manual backup location
    Gosub, resetGUI
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
    FormatTime, sNow, %a_now% T12, (yyyy-MM-dd_HH-mm-ss)
    SplitPath, sPath, dname
    sBackupPath := OutputVar3 . "\STBackup_" . dname . "_" . sNow
    sBackupLogPath := sBackupPath . "\stbackup_log.txt"
    Gui +Disabled
    SB_SetText(A_Tab  . "Backing up...",1,1)
    FileDelete, %sBackupLogPath%
    FileRemoveDir, %sBackupPath%, 1
    If (bCopyallExts)
     {
        ErrorCount := CopyFiles("*",sBackupPath,,bRecursive)
        logErrors("*", sBackupPath, ErrorCount, false)
        If (ErrorCount > 0)
        {   msgBox ,% errIcon,, Backup failed!
            Gosub, resetGUI
            Return
        }
    }
    else
    {
        loop, %ExtArr0%
        {
            if(ExtArr%A_Index% <> "")
            {
                tempExt:=ExtArr%A_Index%
                ErrorCount := CopyFiles(tempExt,sBackupPath,,bRecursive)
                logErrors(tempExt, sBackupPath, ErrorCount, false)
                If (ErrorCount > 0)
                {
                    msgBox ,% errIcon,, Backup failed!
                    Gosub, resetGUI
                    Return
                }
            }
        }
    }
    zipBackup(sBackupPath)
    sCustomDest := OutputVar3
    myINI.iniEdit("History","Last Manual Backup Location",sCustomDest)  
    msgBox ,% infoIcon,, Backup finished.
    Gosub, resetGUI
    ;Run, Explorer /n`,/e`,%sCustomDest%
    return
}

ExitFunc()
{
    Global _bExiting
    if (_bExiting)
    {
        return
    }
    _bExiting := True
    SetTImer, Backup, Off
    sleep, 50
    GuiControlGet, Extstring ,, extsediVar,
    trimExts(Extstring)
    if(Extstring ="")
    {
        sExts := "*;"
    }
    else  
    {
        sExts := Extstring
    }
    if (_sPath<>"" and (util.StrReplaceVar(_sPath) = sPath)) {
        util.setVar(Paths_FilesLocation,_sPath)
    }
    else 
    {
        util.setVar(Paths_FilesLocation,sPath)
    }
    if (_sDest<>"" and (util.StrReplaceVar(_sDest) = sDest)) {
        util.setVar(Paths_BackupsLocation,_sDest)
    }
    else 
    {
        util.setVar(Paths_BackupsLocation,sDest)
    }
    util.setVar(History_LastManualBackupLocation,sCustomDest)
    util.setVar(Option_BackupInterval,tInterval,300000)
    util.setVar(Option_BackupsCount,iBackupCount,10)
    util.setVar(Option_MaxLogSize,iMaxLogSize,500)
    util.setVar(History_NextBackupNumber,iBkupNum)
    util.setVar(Option_Extensions,sExts,"*;")
    util.setVar(Option_Recursive,bRecursive)
    util.setVar(Option_UnlimitedBackups,bInfiniteBkup)
    myINI.iniSave()
    ExitApp
}
