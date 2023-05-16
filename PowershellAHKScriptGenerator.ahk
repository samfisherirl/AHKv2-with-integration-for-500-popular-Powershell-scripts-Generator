#Requires Autohotkey v2.0
#Include JXON.ahk
;AutoGUI 2.5.8
;Auto-GUI-v2 credit to Alguimist autohotkey.com/boards/viewtopic.php?f=64&t=89901
;AHKv2converter credit to github.com/mmikeww/AHK-v2-script-converter
logger := A_ScriptDir "\Script\log.txt"

class parameters { ; original_txt, index, change, variable
	__New(original_txt, index, change, variable := "", int := False) {
		;original_txt, index, change, variable
		this.original_txt := original_txt
		; text in-between dollar signs; suchas
		;   param([string]$PathToExecutables = "")
		;   param([string]
		;   PathToExecutables = "")
		this.index := index
		this.change := change
		this.updated_value := ""
		this.variable := Trim(variable)
		this.int := int
	}
}

class varMenu
{
	__New(index) {
		this.gui := False
	}
	secondGUI() {
		global list_of_params, values := []
		this.gui := Gui()
		; this.gui.Opt("+Owner" MyGui.Hwnd)
		this.gui.Opt("+ToolWindow")
		this.gui.Opt("+AlwaysOnTop")
		
		for i in list_of_params {
			if i.change == True {
				if i.int == False {
				this.gui.Add("Text", , i.variable)
				values.Push(this.gui.Add("Edit"))
				}
				else if i.int == True {
					this.gui.Add("Text", , i.variable)
					values.Push(this.gui.Add("Edit", "Number"))
				}
			}
		}
		but := this.gui.Add("Button", , "Submit")
		but.OnEvent("Click", this.close)
		this.gui.Show()
		Return
	}
	close(*) {
		global list_of_params, values, status
		x := 1
		for i in list_of_params {
			if (i.change == True) {
				i.update_value := values[x].Value
				x := x + 1

			}
		}
		status := 1
		this.gui.Destroy()
	}
}

config_path := A_ScriptDir "\Scripts\config.json"
; config(config_path)
global params := Map()
myGui := Gui()
global Vars := []
global LV := myGui.Add("ListView", "x18 y9 w562 h180", ["Script", "Description"])
LV.OnEvent("DoubleClick", LV_DoubleClick)
import_ps_paths_and_descriptions()
Script_Storage := myGui.Add("Edit", "x18 y210 w560 h180")
MakeButton := myGui.Add("Button", "x352 y446 w120 h33", "Make AHK")
ogcButtonRun := myGui.Add("Button", "x18 y446 w120 h33", "Run Now")
MakeButton.SetFont("s14", "Arial")
MakeButton.OnEvent("Click", MakeAHK)
ogcButtonRun.OnEvent("Click", RunNow)
ogcButtonRun.SetFont("s14", "Arial")
myGui.Title := "Powershell"
LV.ModifyCol()
myGUI.Show("w600 h500")
Return


import_ps_paths_and_descriptions() {
	global LV
	config := Map()
	; IF param([string]$URL = "")
	if FileExist(config_path) {
		try {
			readJSONmapConfig()
		} catch {
			newconfig()
		}
	}
	else {
		newconfig()
	}
}

readJSONmapConfig() {
	global LV
	x := FileRead(config_path)
	config := Jxon_Load(&x)
	for key, value in config {
		LV.Add(, key, value)
	}
}

newconfig() {
	config := Map()
	Loop Files, A_ScriptDir "\Scripts\*.txt" {
		x := FileRead(A_LoopFilePath)
		descr := getDescr(x)
		config.Set(A_LoopFileName, descr)
	}
	FileAppend(Jxon_Dump(config), config_path)
	readJSONmapConfig()
}

RunNow(*) {
	global Vars, LV
	newPS := parsePSvariables(Script_Storage.Value)
	MsgBox(newPS)
	if FileExist(logger) {
		FileMove(logger, A_ScriptDir "\Script\trash.txt", 1)
	}
	FileAppend(newPS, logger)
	Run("powershell -noexit -ExecutionPolicy Bypass -Command `"& { $scriptContent = Get-Content -Path 'log.txt' -Raw; Invoke-Expression -Command $scriptContent }`"")

	Run("Powershell.exe -noexit -Command" Trim(newPS))
	; RowNum := LV.GetNext()  ; Get the text from the row's first field.
	; text := LV.GetText(RowNum)
	; Script_Storage.Value := LTrim(FileRead(A_ScriptDir "\Scripts\" text))
}

MakeAHK(*) {
	global Vars, LV
	newPS := parsePSvariables(Script_Storage.Value)
	if FileExist(logger) {
		FileMove(logger, A_ScriptDir "\Script\trash.txt", 1)
	}
	ahkPath := FileSelect("S", A_ScriptDir, , "AHK (*.ahk)")
	runner := "
	(
		ps_path := A_ScriptDir "\ps_script.txt"
		FileAppend(PS_Script(ps_path), ps_path)
		RunWait(command())
		FileDelete(ps_path)

		command() {
		return "powershell -noexit -ExecutionPolicy Bypass -Command ``"& { $scriptContent = Get-Content -Path 'ps_script.txt' -Raw; Invoke-Expression -Command $scriptContent }``""
		}
		
		PS_Script(ps_path) {  
			if FileExist(ps_path) {
				FileDelete(ps_path)
			}
			
	)"
	if not InStr(ahkPath, ".ahk") {
		ahkPath := ahkPath . ".ahk"
	}
	runner2 := "`n`treturn `"`n`t(`n`t" newPS "`n`t)`"`n}`n"
	if FileExist(ahkPath){
		FileDelete(ahkPath)
	}
	FileAppend(runner . runner2, ahkPath)
	; RowNum := LV.GetNext()  ; Get the text from the row's first field.
	; text := LV.GetText(RowNum)
	; Script_Storage.Value := LTrim(FileRead(A_ScriptDir "\Scripts\" text))
}


parsePSvariables(txt) {
	global params := Map()
	global status := 0
	global list_of_params := []
	index := 0
	runSecondaryGui := 0
	foundParam := 0
	newPS := ""
	Loop Parse, txt, "`n", "`r" {
		if (Trim(A_LoopField) == "") || InStr(Trim(A_LoopField), "exit") {
			continue
		}
		if InStr(A_LoopField, "param(") && (foundParam == 0) {
			foundParam := 1
			dollarsign := StrSplit(A_LoopField, "$")
			index := dollarsign.Length - 1
			for i in dollarsign {
				RegExMatch(i, "\b\w+\s=\s\d+", &out)
				if InStr(i, "`"`"") {
					variablename := StrSplit(i, "=")[1]
					index := index + 1
					list_of_params.Push(parameters(i, index, True, variablename))
				}
				else if InStr(i, "`"") {
					variablename := StrSplit(i, "=")[1]
					temp := StrSplit(i, "`"")
					try {
						temp := temp[1] "`"`"" temp[3]
					} catch {
						temp := temp[1] "`"`"" temp[2]
						
					}
					index := index + 1
					list_of_params.Push(parameters(temp, index, True, variablename))
				}
				else if out {
					variablename := StrSplit(i, "=")[1]
					index := index + 1
					list_of_params.Push(parameters(i, index, True, variablename, True))
				}
				else {
					index := index + 1
					list_of_params.Push(parameters(i, index, False))
				}
			}
			if runSecondaryGui == 0 {
				runSecondaryGui := 1
				v := varMenu(index)
				v.secondGUI()
				store := ""
				while status == 0 {
					Sleep(10)
					if status == 1 {
						break
					}
				}
			}
			for i in list_of_params {
				if i.change == False {
					store .= i.original_txt
				}
				else if i.change == True {
					if i.int == False {
						x := StrReplace(i.original_txt, "`"`"", "`"" i.update_value "`"")
						store .= "$" x
					}
					if i.int == True {
						tomatch := "(\b\w+\s=\s)\d+"
						toreplace := "$1" "$" . i.update_value
						x := RegExReplace(i.original_txt, tomatch, toreplace)
						store .= "$" x
					}
				}
			}
			newPS .= store "`n`t"
		}
		else {
			newPS .= A_LoopField "`n`t"
		}
	}
	return newPS
}

stripVariables(LoopField) {
	variables := []
	quotes := StrSplit(LoopField, "`"`"")
	Loop quotes.Length - 1 {
		quotes := StrSplit(LoopField, "`"`"")
		variables.Push(StrSplit(quotes[A_Index], "$")[2])
	}
}

getDescr(txt) {
	synop := 0
	Loop Parse, txt, "`n", "`r" {
		if (synop == 1) {
			return Trim(A_LoopField)
		}
		else {
			if InStr(A_LoopField, "SYNOPSIS") {
				synop := 1
				continue
			}
		}
	}
}

LV_DoubleClick(LV, RowNumber)
{
	global Script_Storage
	RowText := LV.GetText(RowNumber)  ; Get the text from the row's first field.
	Script_Storage.Value := LTrim(FileRead(A_ScriptDir "\Scripts\" RowText))

}
GuiEscape(*)
{ ; V1toV2: Added bracket
GuiClose:
	ExitApp()
} ; Added bracket in the end
 