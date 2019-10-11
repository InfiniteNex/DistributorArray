; new Distributor Array using python

#SingleInstance force

; Base variables===============================================================================
{
varCode1 := 
varCode2 := 123
varDistributorLong := 
varDistributorShort := 
varCountry := 
}


; Generate startup GUI=========================================================================
{
Gui, Add, Picture, x0 y0 w315, DataBaseArrayBG.png ; BACKGROUND

Gui, Add, Text, vCode1 x50 y15 w30, %varCode1%
Gui, Add, Text, vCode2 x50 y45 w30, %varCode2%
Gui, Add, Text, vDistL x90 y45 w220, %varDistributorLong%
Gui, Add, Text, vDistS x140 y15 w30, %varDistributorShort%
Gui, Add, Text, vCount x180 y15 w20, %varCountry%


Gui, Add, Button, x3 y5 w40 h30 gUpdate , Search
Gui, Add, Button, x3 y35 w40 h30 gCode2 , Raw
;~ Gui, Add, Button, x95 y30 w220 h30 gLong , Long Name
Gui, Add, Button, x90 y5 w40 h30 gShort , Short Name
;~ Gui, Add, Button, x360 y30 w42 h30 , Country

; open sheet for new brands
Gui, Add, Button, x240 y1 w30 h13 gNewBrands, new
; Drag-move panel
Gui, Add, Picture, x1 y1 w240 h5 gUImove, titlebar.png
; copy OEM to clipboard button
Gui, Add, Button, x270 y1 w30 h13 gOEM, oem
;~ Gui, Add, Button, x300 y10 w13 h13 gUpdate, 
Gui, Add, Button, x300 y1 w13 h13 gClose, x

Gui +LastFound +AlwaysOnTop -Caption +ToolWindow
Gui, Show, x880 y7 w315 h70, Search Distributor
return
}



; Button Labels========================================================================
{
Code1:
Clipboard := varCode1
return
Code2:
Clipboard := varDistributorShort
Run, RawDataDownloader.ahk
return
Long:
Clipboard := varDistributorLong
return
Short:
Clipboard := varDistributorShort
return

UImove:
PostMessage, 0xA1, 2,,, A
return

OEM:
Clipboard := 227637
SplashTextOn, 35, , OEM.
WinMove, OEM, , 1150, 30
Sleep, 1000
SplashTextOff
return

NewBrands:
Run, C:\Users\simeon.todorov\Desktop\Import_Feature_Values_Template.xlsm - Shortcut
return
}


; get row data from python and split it in an array

; Update based on new data==============================================================
{
Update:
Run, DistArray.pyw,,, NewPID
Process, WaitClose, %NewPID%
Clipboard := Trim(Clipboard, " `t`r`n")

1Array := StrSplit(Clipboard, ",", "'" "[" "]" "'")

varCode1 := 1Array[1]
varCode2 := 1Array[2]
varDistributorLong := 1Array[3]
StringTrimLeft, varDistributorLong, varDistributorLong, 2
varDistributorShort := 1Array[4]
StringTrimLeft, varDistributorShort, varDistributorShort, 2
varCountry := 1Array[5]
StringTrimLeft, varCountry, varCountry, 2

; Update the text fields
GuiControl, Text, Code1, %varCode1%
GuiControl, Text, Code2, %varCode2%
GuiControl, Text, DistL, %varDistributorLong%
GuiControl, Text, DistS, %varDistributorShort%
GuiControl, Text, Count, %varCountry%

Clipboard := varDistributorShort

return
}


; Button Label for app exit====================================================================
Close:
ExitApp