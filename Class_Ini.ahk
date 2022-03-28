

Class Ini {
    /*
    iniInit will read the specified INI file and load all section and key names,
    and assign all the associated keys to variables of section_keyname.
    
    iniLoad will redo the iniInit function.
    
    iniSave will save all variables to the correct section & key.
    */
    inifile := ""

    __New(inifile) {
        if inifile {
            this.inifile := inifile
            this.iniInit(inifile)
        } 
        else {
             Throw, "Please specify a file name"
        }
    }

    __Delete() {
      This.Free()
    }

    iniInit(inifile){ 
      global 
      local key,var
      inisections := 0      
      loop, read, %inifile% 
      { 
        if regexmatch(A_Loopreadline,"^\[(.*)?]")
          { 
            inisections += 1 
            section%inisections% := regexreplace(A_loopreadline,"(\[)(.*)?(])","$2") 
            section%inisections%_keys := 0
            CurSectionName := regexreplace(A_loopreadline,"(\[)(.*)?(])","$2")
            AllSections = %AllSections%|%CurSectionName%
          } 
          else if regexmatch(A_LoopReadLine,"(\w+)=(.*)") 
          { 
            section%inisections%_keys += 1 
            key := section%inisections%_keys 
            section%inisections%_key%key% := regexreplace(A_LoopReadLine,"(\w+)=(.*)","$1")
            keyval := StrSplit(A_LoopReadLine, "=", " `t")
            var := StrReplace(CurSectionName . "_"  . keyval[1],A_Space)
            %var% := keyval[2]
          } 
        }
    }
    
    iniLoad(inifile := "") {
      if not inifile {
        inifile:=this.inifile
      }
      this.iniInit(inifile)
    }

    iniSave(inifile := "") {
      global
      local sec, var, str
      str :=  ""
      if not inifile {
        inifile := this.inifile
      }
      loop, %inisections%
        {
          sec := A_index
          str := str . "`n" . "[" . section%sec% . "]"
          loop,% section%a_index%_keys
            {
              var := section%sec% "_" section%sec%_key%A_index%
              str := str . "`n" . SubStr(var,InStr(var,"_") + 1)
              var := StrReplace(var, A_Space)
              var := %var%
              str := str . "=" . var
            }
        }
        str := Trim(str, "`n`r `t") 
        FileDelete, %inifile%
        FileAppend, %str%,%inifile%,UTF-16
    }
    
    iniEdit(sec, key, val, inifile := "") {
        if not inifile {
            inifile := this.inifile
        }
        IniWrite, %val%, %inifile%, %sec%, %key%
        this.iniLoad()
    }
}