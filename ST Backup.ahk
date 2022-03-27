#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent 
#SingleInstance, ignore 
#NoTrayIcon

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

_font:="Tahoma"

getVersion(ver)
{
    index := InStr(ver,".")
    return substr(ver,1,index+1) "." substr(ver,index+2,1)
}

myTitle:= myName " " getVersion(curVersion)

IsEmpty(Dir){
   Loop %Dir%\*.*, 0, 1
      return 0
   return 1
}

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

psEscape(sPath)
{
    return RegExReplace(sPath, "[\[\]]", "``$0")
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
    FileDelete, %sParent%\%sName%.stb.zip
    FileDelete, %sPath%\%sName%.stb.zip
    sOut := sParent "\" sName ".stb.zip"
    psOut := psEscape(sOut)
    psPath := psEscape(sPath)
    Run PowerShell.exe -Command (Compress-Archive -Path '%psPath%\*' -CompressionLevel Optimal -DestinationPath '%sOut%'); if ($?) {(Remove-Item -force '%psPath%' -recurse -Confirm:$False);},, Hide UseErrorLevel
    if (ErrorLevel = "ERROR")
    {
        Zip(sPath , sParent "\" sName ".stb.zip")      
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
        sLog := "*." . sExt . " Backup: """ . trimPath(sBackupPath . _backup_ext) . """"
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
        return 0
    }
    return 1
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

setVar(ByRef var,value,def:="NULL") 
{
    if (value) 
    {
        var:=value
    } 
    else if (def!="NULL") 
    {
      var:=def  
    }
}

CopyFiles(sExt,sDestFolder,sSourceFolder:="",bRecursive:=false)   
{
    global sPath
    global sDest
    sPath:= trimPath(sPath)
    sDest:= trimPath(sDest)
    sDestFolder:= trimPath(sDestFolder)
    if !sSourceFolder
        sSourceFolder:=sPath
    if IsEmpty(sSourceFolder)
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
            ErrCount := ErrorLevel
            If (ErrCount > 0)
            {
                Return ErrCount
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
    sleep, 50
    myINI := new Ini(A_ScriptDir . "\STB_settings.ini")
    setvar(_sPath,Paths_FilesLocation)
    setvar(_sDest,Paths_BackupsLocation)
    setvar(sCustomDest,History_LastManualBackupLocation)
    setvar(tInterval,Option_BackupInterval,300000)
    setvar(iBackupCount,Option_BackupsCount,10)
    setvar(iMaxLogSize,Option_MaxLogSize,500)
    setvar(iBkupNum,History_NextBackupNumber,1)
    setvar(sExts,Option_Extensions,"*;")
    setvar(bRecursive,Option_Recursive)
    setvar(bInfiniteBkup,Option_UnlimitedBackups)
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
    if (sCustomDest<>"" and sCustomDest=sPath) 
    {
        sCustomDest .="\ST_Backups"
        IniWrite, %sCustomDest%, STB_settings.ini, History, Last Manual Backup Location
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
    FileAppend, %str%,%A_ScriptDir%\STB_settings.ini,UTF-16 
    sleep, 50
    myINI := new Ini(A_ScriptDir . "\STB_settings.ini")
}

Hotkey, ^!x, ExitSub
OnExit, ExitSub

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
Gui,Font, s7 , Lucida Console
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
	AutoSize("reset") ; Needs to reset if you changed the Control size manually.
    return

GuiSize:
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
    IniWrite, %sPath%, STB_settings.ini, Paths, Files Location
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
    checkNum(InVar, "BCedit", 1,10000)
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
    
Recursivecbx:
{
    bRecursive := !bRecursive
    IniWrite, %bRecursive%, STB_settings.ini, Option, Recursive
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
    IniWrite, %bInfiniteBkup%, STB_settings.ini, Option, Unlimited Backups
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
    sPath := trimPath(sPath)
    sDest := trimPath(sDest)
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
    IniWrite, %iBkupNum%, STB_settings.ini, History, Next Backup Number
    Return
}

RSbtn:
{
    GuiControlGet, sPath,, SLedit
    GuiControlGet, sDest,, BLedit
    GuiControlGet, Extstring ,, extsediVar,
    sPath := trimPath(sPath)
    sDest := trimPath(sDest)
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
    shortFile := shrinkString(SelectedFile,55,"l")
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
    Unz(SelectedFile, tempPath)
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
    strLog := "Restore: """ . trimPath(SelectedFile) . """"
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
    Run, Explorer /n`,/e`,%sPath%
    return
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
    IniWrite, %sCustomDest%, STB_settings.ini, History, Last Manual Backup Location    
    msgBox ,% infoIcon,, Backup finished.
    Gosub, resetGUI
    Run, Explorer /n`,/e`,%sCustomDest%
    return
}

ExitSub:
{
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
    if (_sPath<>"" and (StrReplaceVar(_sPath) = sPath)) {
        setvar(Paths_FilesLocation,_sPath)
    }
    else 
    {
        setvar(Paths_FilesLocation,sPath)
    }
    if (_sDest<>"" and (StrReplaceVar(_sDest) = sDest)) {
        setvar(Paths_BackupsLocation,_sDest)
    }
    else 
    {
        setvar(Paths_BackupsLocation,sDest)
    }
    setvar(History_LastManualBackupLocation,sCustomDest)
    setvar(Option_BackupInterval,tInterval,300000)
    setvar(Option_BackupsCount,iBackupCount,10)
    setvar(Option_MaxLogSize,iMaxLogSize,500)
    setvar(History_NextBackupNumber,iBkupNum)
    setvar(Option_Extensions,sExts,"*;")
    setvar(Option_Recursive,bRecursive)
    setvar(Option_UnlimitedBackups,bInfiniteBkup)
    myINI.iniSave(A_ScriptDir . "\STB_settings.ini")
    ExitApp
}