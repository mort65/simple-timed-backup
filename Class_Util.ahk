Class Util {
    ; Class initialization
    Static InitClass := Util.ClassInit()
    __New() { ; You must not instantiate this class!
        If (This.InitClass == "!DONE!") { ; external call after class initialization
         This["!Access_Denied!"] := True
         Return False
        }
    }
    ; ----------------------------------------------------------------------------------------------------------------
    __Delete() {
        If This["!Access_Denied!"]
         Return
        This.Free()
    }
    ; ===================================================================================================================
    ; ClassInit       Internal creation of a new instance to ensure that __Delete() will be called.
    ; ===================================================================================================================
    ClassInit() {
        Util := New Util
        Return "!DONE!"
    }
    
    getVersion(ver:="1.000")
    { ; 1.000->1.0.0.0
        index := InStr(ver,".")
        subVer := substr(ver,index+1)
         version := substr(ver,1,index-1)
        Loop, Parse, subVer
        {
            if (A_LoopField<>".") 
            {
                version := version . "." . A_LoopField    
            }     
        }
        return version
    }
    
    IsEmpty(Dir)
    { ; check if a directory has no files
       Loop %Dir%\*.*, 0, 1
          return 0
       return 1
    }
    
    Zip(sDir, sZip)
    { ; compress as zip
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
    { ; extract zip archive
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
    { ; compress with powershell
        RunWait PowerShell.exe -Command Compress-Archive -LiteralPath '%inPath%' -CompressionLevel Optimal -DestinationPath '%outPath%',, Hide UseErrorLevel
        Return ErrorLevel
    }
    psUnzip(inPath,outPath)
    { ; extract with powershell
        RunWait PowerShell.exe -Command Expand-Archive -LiteralPath '%inPath%' -DestinationPath '%outPath%',, Hide UseErrorLevel
        Return ErrorLevel
    }

    psEscape(sPath)
    { ; replace [ and ] with ``[ and ``] in a path string before sending it to powershell.
        return RegExReplace(sPath, "[\[\]]", "``$0")
    }
    
    trimPath(strPath)
    { ; remove extra spaces from begining & end, 
      ; and \ from the end of a path string.
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
    { ; if length of a string is higher than maxlength, 
      ; remove extra characters from one side of the sring (left,right,middle) 
      ; and add '...' to the string.
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
    { ;check if a folder is inside another folder
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

    varExist(ByRef v) 
    { ; Requires 1.0.46+
        return &v = &n ? 0 : v = "" ? 2 : 1
    }

    setVar(ByRef var,value,def:="!NULL!") 
    { ; set value of var to the value of another variable if that variable exists
      ; otherwise if an optional default value is provided, set it to that value.
        var := this.varExist(value) ? value : def == "!NULL!" ? var : def
    }
}