#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=21.ico
#AutoIt3Wrapper_Outfile_x64=EASY_ALL.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=Adds staff printers - Main/Middletown
#AutoIt3Wrapper_Res_Fileversion=2.7.0.0
#AutoIt3Wrapper_Res_ProductVersion=2.7.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Carmine
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=ProductName|Easy Button
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <GUIConstantsEx.au3>
#include <File.au3>
#include <Constants.au3>
#include <Array.au3>
#include <String.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstants.au3>
#include <WindowsConstants.au3>
#include <TrayConstants.au3>
#include <Timers.au3>
#include <Misc.au3>
#include <ColorConstantS.au3>


If _Singleton("EASY", 2) = 0 Then
	MsgBox($MB_SYSTEMMODAL, "Warning", "An occurence of EASY is already running")
	Exit
EndIf

Global $test, $sComboRead, $currentlist

Local $idCheckbox, $cmd3, $cmd4
$idfile = @TempDir & '\setdef.vbs'
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
TrayCreateItem("About")
TrayItemSetOnEvent(-1, "About")
TrayCreateItem("") ; Create a separator line.
TrayCreateItem("Exit")
TrayItemSetOnEvent(-1, "ExitScript")
TraySetOnEvent($TRAY_EVENT_PRIMARYDOUBLE, "About") ; Display the About MsgBox when the tray icon is double clicked on with the primary mouse button.
TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.
Func About()
	; Display a message box about the AutoIt version and installation path of the AutoIt executable.
	MsgBox(262176, "Easy - Add a printer", "" & @CRLF & _
			"Add a printer easily by Carmine Version 2.7.0.0", 4) ; Find the folder of a full path.
EndFunc   ;==>About

Func ExitScript()
	FileDelete(@TempDir & '\red.bmp')
	FileDelete(@TempDir & '\green.bmp')
	FileDelete(@TempDir & '\t.mp3')
	FileDelete($idfile)
	Exit
EndFunc   ;==>ExitScript

FileInstall('red.bmp', @TempDir & '\red.bmp')
FileInstall('green.bmp', @TempDir & '\green.bmp')
;FileInstall('t.mp3', @TempDir & '\t.mp3')
; Create a GUI with various controls.
Local $hGUI = GUICreate("Add Printer rev 3", 260, 110, -1, -1, -1, $WS_EX_TOPMOST)
GUISetBkColor(0x33332F)
; Create a combobox control.
GUISetFont(9, 500, 1, 'Tahoma')
$idComboBox = GUICtrlCreateCombo("", 10, 10, 130, 20)
;GUISetControlFont(-1,12,600,"Times New Roman")
$idCheckbox = GUICtrlCreateCheckbox('', 147, 12, 20, 20)
$idClose11 = GUICtrlCreateButton("HARD", 165, 75, 85, 25, $BS_BITMAP)
GUICtrlSetImage($idClose11, @TempDir & "\red.bmp")
$idADD = GUICtrlCreateButton("EASY", 10, 75, 85, 25, $BS_BITMAP)
GUICtrlSetImage($idADD, @TempDir & "\green.bmp")
$font = "Tahoma" ; sets the $font to tahoma ms
GUISetFont(9, 400, 1, $font) ; sets the font to 9 pt, normal  weight, non italic
GUICtrlSetDefColor(0xFFFFFF)
GUICtrlCreateLabel("Carmine's super EASY add a printer tool", 10, 45)
$idLabel = GUICtrlCreateLabel('Make Default', 170, 14, 80, 20)
GUICtrlSetState($idCheckbox, 4)

Local $we, $MyControl, $verify, $rt, $ii, $FileList
Dim $FileList[1][2]

; http://searchwindowsserver.techtarget.com/tip/List-all-your-network-printers-to-a-file

$zy = FileReadToArray('Printservers.ini')
For $h = 0 To UBound($zy) - 1
	$y = $zy[$h]
	If $y = "" Then
		ContinueLoop
	Else
		$cmd = 'net view ' & $y & ' | find /i "print"'
		Call('qw')
	EndIf
Next

Func qw()
	$we = (_GetDOSOutput($cmd) & @CRLF)
	$aArray1 = _StringExplode($we, @CRLF, 0)
	For $i = 0 To UBound($aArray1) - 1
		If StringInStr($aArray1[$i], "print") > 1 Then
			$a = StringSplit($aArray1[$i], " ", 1)
			_ArrayAdd($FileList, $a[1])
			$t = UBound($FileList) - 1
			$FileList[$t][1] = '\\' & $y & '\'
		Else
			ContinueLoop
		EndIf
	Next
EndFunc   ;==>qw

_ArraySort($FileList, 0)

$cData = ""
For $i = 1 To UBound($FileList) - 1
	$cData &= "|" & $FileList[$i][0]
Next
;or GUICtrlSetData($MyControl, $cData)
GUICtrlSetData($idComboBox, $cData, $FileList[1][0])

Local $sComboRead = ""
; Display the GUI.
GUISetState(@SW_SHOW, $hGUI)

; Loop until the user exits.
Local $hStarttime = _Timer_Init()
While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			FileDelete(@TempDir & '\red.bmp')
			FileDelete(@TempDir & '\green.bmp')
			;FileDelete(@TempDir & '\t.mp3')
			ExitLoop
		Case $idClose11
			GUISetState(@SW_DISABLE, $hGUI)
			;SoundSetWaveVolume(85)
			;SoundPlay(@TempDir & '\t.mp3', 0)
			MsgBox(262160, "Grasshopper says:", "I'm sorry the printer mapping tool did not work for you." & @CRLF & @CRLF & 'You can also try mapping the printer manually by clicking on the printer in the print server after you click OK', 30)
			$path = '\\your-print-server\'; your print server goes here
			Run("explorer /e, " & '"' & $path & '"')
			GUISetState(@SW_ENABLE, $hGUI)
			FileDelete(@TempDir & '\red.bmp')
			FileDelete(@TempDir & '\green.bmp')
			FileDelete(@TempDir & '\t.mp3')
			ExitLoop
		Case $idADD
			$sComboRead = GUICtrlRead($idComboBox)
			$sComboRead = StringStripWS($sComboRead, 3) ; added 3/29/2017
			Local $hStarttime = _Timer_Init()
			$verify = 0
			$cmd4 = 'null'
			Call('verify')
			If $verify = 1 Then
				ContinueCase
			EndIf
			Call('_isinstalled')
			Call('box')
			;$cmd3 = 'rundll32 printui.dll,PrintUIEntry /q /in /n"\\csc-mc-prtstf-1\' & $sComboRead & '"'
			GUISetState(@SW_HIDE, $hGUI)
			If $cmd4 <> 'null' Then
				;MsgBox(0, "", $cmd4)
				RunWait('"' & @ComSpec & '" /c ' & $cmd4, @SystemDir, @SW_HIDE) ; removing for remapping purposes
			EndIf
			Sleep(2000)
			RunWait('"' & @ComSpec & '" /c ' & $cmd3, @SystemDir, @SW_HIDE) ; mapping
			Local $hStarttime = _Timer_Init()
			If _IsChecked($idCheckbox) Then
				Call('def')
			EndIf
			MsgBox(262208, "Grasshopper Says:", '"' & $sComboRead & '"' & " " & "has been successfully mapped." & @CRLF & "Have a Nice Day", 10)
			Local $sComboRead = ""
			GUISetState(@SW_SHOW, $hGUI)
	EndSwitch
	If _Timer_Diff($hStarttime) >= 300000 Then
		FileDelete(@TempDir & '\red.bmp')
		FileDelete(@TempDir & '\green.bmp')
		FileDelete(@TempDir & '\t.mp3')
		Exit
	EndIf
WEnd

; Delete the previous GUI and all controls.
GUIDelete($hGUI)

Func _GetDOSOutput($sCommand)
	Local $iPID, $sOutput = ""
	$iPID = Run('"' & @ComSpec & '" /c ' & $sCommand, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	While 1
		$sOutput &= StdoutRead($iPID, False, False)
		If @error Then
			ExitLoop
		EndIf
		Sleep(10)
	WEnd
	Return $sOutput
EndFunc   ;==>_GetDOSOutput

Func _IsChecked(Const $iControlID)
	Return BitAND(GUICtrlRead($iControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

Func verify()
	If $sComboRead = "" Then
		GUISetState(@SW_DISABLE, $hGUI)
		MsgBox(262160, "Grasshopper says:", "The printer field cannot be left blank." & @CRLF & "Please try again.", 10)
		Local $sComboRead = ""
		GUICtrlSetData($idComboBox, $cData, $FileList[1][0]) ; added 3/29/2017
		GUICtrlSetState($idCheckbox, 4) ; added 4/4/2017
		GUISetState(@SW_ENABLE, $hGUI)
		$verify = 1
		Return
	EndIf

	For $k = 0 To UBound($FileList) - 1
		If $FileList[$k][0] = $sComboRead Then
			Return
		ElseIf $k = UBound($FileList) - 1 Then
			GUISetState(@SW_DISABLE, $hGUI)
			MsgBox(262160, "Grasshopper says:", "The printer " & '"' & $sComboRead & '"' & " is not located on this server." & @CRLF & "Please try again.", 10)
			Local $sComboRead = ""
			GUICtrlSetData($idComboBox, $cData, $FileList[1][0]) ; added 3/29/2017
			GUICtrlSetState($idCheckbox, 4) ; added 4/4/2017
			GUISetState(@SW_ENABLE, $hGUI)
			$verify = 1
		EndIf
	Next
EndFunc   ;==>verify

Func _isinstalled()
	; below checks for currently mapped network printers

	Dim $currentlist[1]
	; Script Start - Add your code below here
	$wbemFlagReturnImmediately = 0x10
	$wbemFlagForwardOnly = 0x20
	$colItems = ""
	$Output = ""
	$nOutput = ""
	$objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_Printer", "WQL", _
			$wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If IsObj($colItems) Then
		For $objItem In $colItems
			If StringInStr($objItem.DeviceID, "\\") Then
				$nOutput = $objItem.DeviceID
				;MsgBox(0,"",$nOutput)
			Else
				$Output = $Output & $objItem.DeviceID & @CRLF
			EndIf
			_ArrayAdd($currentlist, $nOutput)
			$nOutput = ""
		Next

		;FileWrite($fp & "\_NetworkPrinters.cmd", $nOutput)
		;FileWrite($fp & "\_LocalPrinters.cmd", $Output)
	Else
		;FileWrite($fp & "\Printers.txt", "No printers Found")
	EndIf
	;FileSetAttrib($cmdfile, "+R-A")
	For $i = UBound($currentlist) - 1 To 1 Step -1
		If $currentlist[$i] = "" Then
			_ArrayDelete($currentlist, $i)
		EndIf
	Next
	;_ArrayDisplay($w)
	;_ArrayDisplay($currentlist)

	For $i = 1 To UBound($currentlist) - 1
		If StringInStr($currentlist[$i], $sComboRead) > 1 Then
			$cmd4 = 'rundll32 printui.dll,PrintUIEntry /q /dn /n' & '"' & $currentlist[$i] & '"' ; removing for re-,apping purposes.
		EndIf
	Next

EndFunc   ;==>_isinstalled


Func def() ; this creates a vbs script that sets the default printer if the tickbox is selected.
	Sleep(2000)
	FileWrite($idfile, 'Option Explicit' & @CRLF)
	FileWrite($idfile, 'on error resume next' & @CRLF)
	FileWrite($idfile, 'Dim objPrinter, printerPath' & @CRLF)
	FileWrite($idfile, 'printerPath =' & '"' & $rt & $sComboRead & '"' & @CRLF)
	FileWrite($idfile, 'Set objPrinter = CreateObject("WScript.Network")' & @CRLF)
	FileWrite($idfile, 'objPrinter.SetDefaultPrinter printerPath' & @CRLF)
	FileWrite($idfile, 'wscript.sleep 500' & @CRLF)
	FileWrite($idfile, 'WScript.Quit' & @CRLF)
	ShellExecuteWait('setdef.vbs', "", @TempDir)
	GUICtrlSetState($idCheckbox, 4)
	FileDelete($idfile)
EndFunc   ;==>def

Func box() ;
	For $ii = 1 To UBound($FileList) - 1
		If $sComboRead = $FileList[$ii][0] Then
			$rt = $FileList[$ii][1]
			$cmd3 = 'rundll32 printui.dll,PrintUIEntry /q /in /n' & '"' & $FileList[$ii][1] & $sComboRead & '"' ; mapping
		EndIf
	Next
EndFunc   ;==>box

;https://www.autoitscript.com/forum/topic/81809-populating-a-combobox-from-an-array-solved/
;https://community.spiceworks.com/scripts/show/71-install-or-remove-network-printers-and-drivers
;https://blogs.technet.microsoft.com/yongrhee/2009/06/29/how-to-add-printers-with-no-user-interaction-on-windows-vistawindows-server-2008/
;https://community.spiceworks.com/scripts/show/553-setdefaultprinter
