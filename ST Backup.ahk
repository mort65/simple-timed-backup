#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent 
#SingleInstance, ignore 
sPath := ""
sDest := ""
_sPath:=""
_sDest:=""
sCustomDest := ""
sExts :=""
sBackupt := "Backup is Running!"
sBackupf := "Backup is Stopped!"
iBackupCount := 10
tInterval := 300000 ; 5 min
toggle := 0
sCurrentTime :=""
bIsEDExtsenabled:=-1
bCopyallExts:=false
bRecursive:=false
red:="c0xe1256b"
blue:="c0x056bed"
bZipBackup := 0
errIcon := 16
infoIcon := 64
curVersion:=1.0
myName:="Simple Timed Backup"
myTitle:=myName . " " . curVersion

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
        SB_SetText(A_Tab curTime . " Backup: """ . trimPath(shrinkString(sBackupPath,70,"m")) . """",1,1)
    } else {
        if FileExist(sMainLogPath)
        {
            FileAppend ,`n%sNow% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            FileAppend ,`n%sNow% can`t copy %errCount% file(s)!
        }else  {
            FileAppend ,%sNow% warning! `, extension:%sExt% `, source:%sPath%\ `, destination:%sBackupPath%\, %sMainLogPath%
            FileAppend ,`n%sNow% can`t copy %errCount% file(s)!
        }
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
{   break:=false
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
    sPath := StrReplaceVar(_sPath)
    sDest := StrReplaceVar(_sDest)
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
}

Hotkey, ^!x, ExitSub
OnExit, ExitSub

Gui +LastFound
;Gui, -theme
;WinSet, Transparent, 254
GUI, -ToolWindow
Gui, +CAPTION
GUI, -MaximizeBox
Gui, Margin, 0, 0
Gui,Font, 
;Gui, Add, GroupBox, x8 y4 w486 h466,
gui,font,italic s8
Gui,Add,Edit,x110 y46 w260 h22 %black%  ReadOnly  vSLedit,
GuiControl,, SLedit, %sPath%
Gui,Add,Edit,x110 y118 w260 h22 %black% ReadOnly  vBLedit,
gui,font,
GuiControl,, BLedit, %sDest%
Gui,Add,Button,x378 y46 w50 h22 vSPvar gSPbtn,Change
Gui,Add,Button,x378 y118 w50 h22  vBPvar gBPbtn,Change
Gui, Add, Button, x436 y46 w50 h22  vOSPvar gOSPbtn, Open
Gui, Add, Button, x436 y118 w50 h22  vOBPvar gOBPbtn, Open
Gui,Add,Text,x16 y48 w90 h13 %black% left,Files to Backup:
Gui,Add,Text,x16 y120 w90 h13 %black% left ,Backups Location:
Gui,Add,Edit,x120 y188 w38 h24 %black% Number ReadOnly  vBIedit gBIedit
mInterval := (tInterval/60000)
Gui,Add,UpDown, 0x20  Range1-720 ,%mInterval%,vBIud
Gui,Add,Edit,x120 y224 w38 h24 %black% Number ReadOnly  vBCedit gBCedit
Gui,Add,UpDown, 0x20  Range1-100,%iBackupCount%,vBCud
Gui,Add,Button,x80 y410 w110 h40 center vACvar gACbtn,Activate
Gui,Add,Button,x302 y410 w110 h40 +Disabled vDEvar gDEbtn,Deactivate
Gui,Add,Text,x33 y192 w80 h13 %black% left ,Backup every:
Gui,Add,Text,x164 y192 w40 h25 %black% ,minutes
Gui,Add,Text,x33 y228 w80 h13 %black% left ,Backups count:
Gui,Font,Normal s14  Bold ,Segoe UI
Gui,Add,Text,x30 y348 w200 h50 Center %red% vNotetext,%sBackupf%
Gui,Font,Normal s10
Gui,Add,Edit,x265 y200 w185 h103 %black% r4 1024 Lowercase Multi Border readonly 64 vextsediVar gextsEdit,%sExts%
Gui,Font,
Gui,Add,Text,x268 y175 w140 h20 %black% -Wrap,File extensions to backup:
Gui,Add,Button,x335 y290 w45 h25  vEDbtnvar gextsEDbtn,Edit
Gui,Add,Button,x270 y290 w45 h25 disabled  vEDbtnokvar gextsEDokbtn,Ok
Gui,Add,Button,x400 y290 w45 h25 disabled  vEDbtncancelvar gextsEDcancelbtn,Cancel
Gui,Add,Button, x322 y344 w70 h34 center +Disabled  vBKvar gBKbtn , Manual Backup

if sPath !=
    GuiControl, Enabled, BKvar

Gui,Add,Checkbox,x33 y264 w100 h20 %black% -Wrap  vZipBackupvar gZipBackupcbx,Zip backups?

if (bZipBackup = 1)
{
    GuiControl,, ZipBackupvar, 1    
}
else
{
    GuiControl,, ZipBackupvar, 0
}

Gui,Add,Checkbox,x33 y300 w100 h20 %black% -Wrap  vRecursiveVar gRecursivecbx,Recursive?

if (bRecursive = 1)
{
    GuiControl,, RecursiveVar, 1    
}
else
{
    GuiControl,, RecursiveVar, 0
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
BIedit_TT := "Automated backups will be created at the selected interval."
EDbtnvar_TT := "Edit what file types to backup."
;Gui, -theme
;Gui,Font,Normal s11
Gui, Add, StatusBar,,
Gui,Font,
Gui,Show,x390 y122 w500 h500 ,%myTitle%
OnMessage(0x200, "WM_MOUSEMOVE")
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

extsEDcancelbtn:
{
    GuiControl,,extsediVar, %sExts%
    GuiControl, +ReadOnly, extsediVar
    GuiControl, Disabled, EDbtncancelvar
    GuiControl, Disabled, EDbtnokvar
    GuiControl, Enabled, EDbtnvar
    bIsEDExtsenabled := bIsEDExtsenabled * -1
    Return
}
    
extsEDbtn:
{
    if(bIsEDExtsenabled = -1)
    {
        GuiControl, -ReadOnly, extsediVar
        GuiControl, Enabled, EDbtncancelvar
        GuiControl, Enabled, EDbtnokvar
        GuiControl, Disabled, EDbtnvar
        bIsEDExtsenabled := bIsEDExtsenabled * -1
        return
    }else  {
        GuiControl, +ReadOnly, extsediVar
        GuiControl, Disabled, EDbtncancelvar
        GuiControl, Disabled, EDbtnokvar
        GuiControl, Enabled, EDbtnvar
        bIsEDExtsenabled := bIsEDExtsenabled * -1
        return
    }
}

extsEDokbtn:
{
    GuiControlGet, Extstring ,, extsediVar,
    trimExts(Extstring)
    sExts := Extstring
    If InStr(sExts, "*")
        sExts := "*;"
    if sExts =
        sExts := "*;"
    GuiControl,,extsediVar, %sExts%
    GuiControl, +ReadOnly, extsediVar
    GuiControl, Disabled, EDbtncancelvar
    GuiControl, Disabled, EDbtnokvar
    GuiControl, Enabled, EDbtnvar
    bIsEDExtsenabled := bIsEDExtsenabled * -1
    Return
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

OSPbtn:
{
    if (sPath="")
        Return
    if InStr(FileExist(sPath),"D")
    {
        Run, Explorer /n`,/e`,%sPath%
    }
    else 
    {
        SplashTextOff
        msgbox,% errIcon,, The path you entered could not be found: %sPath%
    }
    return
}

OBPbtn:
{
    if (sDest="")
        Return
    if InStr(FileExist(sDest),"D")
    {
        Run, Explorer /n`,/e`,%sDest%
    }
    else
    {
        SplashTextOff
        msgbox,% errIcon,, The path you entered could not be found: %sDest%
    }
    return
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
    if(bIsEDExtsenabled = 1)
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
        SplashTextOff
        msgbox,% errIcon,, The path you entered could not be found: %sPath%
        return
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
            sDest:=sPath
            sDest.="\ST_Backups"
            GuiControl,, BLedit, %sDest%
            IniWrite, %sDest%, STB_settings.ini, Paths, Backups Location
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
        GuiControl, Disabled, EDbtncancelvar
        GuiControl, Disabled, EDbtnokvar
        GuiControl, Disabled, ZipBackupvar
        GuiControl, Disabled, RecursiveVar
        GuiControl,,Notetext,%sBackupt%
        Gui,Font,Normal s14 Bold %blue% ,Segoe UI
        GuiControl, Font, Notetext 
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
    if(bIsEDExtsenabled = -1)
    {
        GuiControl,Enable,EDbtnvar
    }else  {
        GuiControl, Enable, EDbtncancelvar
        GuiControl, Enable, EDbtnokvar
    }
    GuiControl,,Notetext,%sBackupf%
    Gui,Font,Normal s14 Bold %red% ,Segoe UI
    GuiControl, Font, Notetext 
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
        FormatTime, sNow, %a_now% T12, [yyyy-MM-dd%a_space%HH:mm:ss]
        FileAppend ,`n%sNow% exiting program..., %sLogfullpath%
        sleep, 50
    }
    ExitApp
}