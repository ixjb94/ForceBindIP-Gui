#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

#include <Constants.au3>

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

#Region ### START Koda GUI section ### Form=C:\Program Files (x86)\AutoIt3\koda_1.7.3.0\Extras\Import\ForceBindIP.kxf
$ForceBindIP = GUICreate("ForceBindIP - By IX JB", 627, 429, -1, -1)
$Group_appAddress = GUICtrlCreateGroup("Application Address", 16, 16, 593, 153)
$Button_runApp = GUICtrlCreateButton("Run Application", 240, 120, 123, 25)
$Input_appAddress = GUICtrlCreateInput("Application Address", 32, 72, 553, 21)
$Combo_internetSelect = GUICtrlCreateCombo("Select Internet Connection", 32, 40, 553, 25)
$Radio_x86 = GUICtrlCreateRadio("x86", 440, 112, 113, 17)
$Radio_x64 = GUICtrlCreateRadio("x64", 440, 136, 113, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group_Quick = GUICtrlCreateGroup("Quick Access", 16, 176, 593, 209)
$Button1 = GUICtrlCreateButton("", 40, 200, 91, 25)
$Button2 = GUICtrlCreateButton("", 192, 200, 91, 25)
$Button3 = GUICtrlCreateButton("", 352, 200, 91, 25)
$Button4 = GUICtrlCreateButton("", 496, 200, 91, 25)
$Button5 = GUICtrlCreateButton("", 40, 240, 91, 25)
$Button6 = GUICtrlCreateButton("", 192, 240, 91, 25)
$Button7 = GUICtrlCreateButton("", 352, 240, 91, 25)
$Button8 = GUICtrlCreateButton("", 496, 240, 91, 25)
$Button9 = GUICtrlCreateButton("", 40, 280, 91, 25)
$Button10 = GUICtrlCreateButton("", 192, 280, 91, 25)
$Button11 = GUICtrlCreateButton("", 352, 280, 91, 25)
$Button12 = GUICtrlCreateButton("", 496, 280, 91, 25)
$Button_Config = GUICtrlCreateButton("Config", 432, 344, 155, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Label_copyRight = GUICtrlCreateLabel("IX JB", 520, 400, 84, 17)
$Label_Internet = GUICtrlCreateLabel("Select Internet Connection", 16, 400, 474, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


; Vars
Global $cFilePath = @ScriptDir & "\ForceBindIP-config.ini"

If FileExists($cFilePath) = False Then
	createIniFile() ; Create config.ini
EndIf

GetInternet()
setIniName()

While 1
   $nMsg = GUIGetMsg()

	Switch $nMsg

	  Case $GUI_EVENT_CLOSE
			Exit

	  Case $Button_runApp
		 If GUICtrlRead($Input_appAddress) <> "" Then

			$x = ""
			If (GUICtrlRead($Radio_x86) == 1) Then
				$x = "86"
			ElseIf (GUICtrlRead($Radio_x64) == 1) Then
				$x = "64"
			EndIf

			runCMD(GUICtrlRead($Input_appAddress), $x)
		 Else
			MsgBox(0,"Input Can Not Be Empty","Please Enter Valid Data...")
		 EndIf

	  Case $Button1
		 runCMD(getIniAddress("Button1"), getIniX("Button1"))

	  Case $Button2
		 runCMD(getIniAddress("Button2"), getIniX("Button2"))

	  Case $Button3
		 runCMD(getIniAddress("Button3"), getIniX("Button3"))

	  Case $Button4
		 runCMD(getIniAddress("Button4"), getIniX("Button4"))

	  Case $Button5
		 runCMD(getIniAddress("Button5"), getIniX("Button5"))

	  Case $Button6
		 runCMD(getIniAddress("Button6"), getIniX("Button6"))

	  Case $Button7
		 runCMD(getIniAddress("Button7"), getIniX("Button7"))

	  Case $Button8
		 runCMD(getIniAddress("Button8"), getIniX("Button8"))

	  Case $Button9
		 runCMD(getIniAddress("Button9"), getIniX("Button9"))

	  Case $Button10
		 runCMD(getIniAddress("Button10"), getIniX("Button10"))

	  Case $Button11
		 runCMD(getIniAddress("Button11"), getIniX("Button11"))

	  Case $Button12
		 runCMD(getIniAddress("Button12"), getIniX("Button12"))

	  Case $Button_Config
		  ShellExecute($cFilePath)

	  Case $Combo_internetSelect
		 $ComboString = GUICtrlRead($Combo_internetSelect)
		 $iPosition = StringInStr($ComboString, "{")
		 $iPosition = $iPosition - 1
		 $sString = StringTrimLeft($ComboString,$iPosition)
		 GUICtrlSetData($Label_Internet,$sString)

	EndSwitch
 WEnd




Func getIniName($section,$param)

	; Ini Array
	Local $iArray = IniReadSection($cFilePath, $section)
	$name = $iArray[1][1] ; Name
	GUICtrlSetData($param,$name)

EndFunc

Func setIniName()

	getIniName("Button1",$Button1)
	getIniName("Button2",$Button2)
	getIniName("Button3",$Button3)
	getIniName("Button4",$Button4)
	getIniName("Button5",$Button5)
	getIniName("Button6",$Button6)
	getIniName("Button7",$Button7)
	getIniName("Button8",$Button8)
	getIniName("Button9",$Button9)
	getIniName("Button10",$Button10)
	getIniName("Button11",$Button11)
	getIniName("Button12",$Button12)

EndFunc


Func getIniAddress($section)
	; Ini Array
	Local $iArray = IniReadSection($cFilePath, $section)
	$address = $iArray[2][1] ; Address
	Return $address

EndFunc

Func getIniX($section)
	; Ini Array
	Local $iArray = IniReadSection($cFilePath, $section)
	$x = $iArray[3][1] ; X
	Return $x
EndFunc


Func createIniFile()

	For $i = 1 To 12 Step 1
		IniWrite ($cFilePath , "Button" & $i , "Name", "Button" & $i)
		IniWrite ($cFilePath , "Button" & $i , "Address", "")
		IniWrite ($cFilePath , "Button" & $i , "X", "" & @CRLF)
	Next

EndFunc


Func GetInternet ()

	; Get WMI
	$wbemFlagReturnImmediately = 0x10
	$wbemFlagForwardOnly = 0x20
	$colItems = ""
	$strComputer = "localhost"
	$Output=""
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID != NULL", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	If IsObj($colItems) then
	  For $objItem In $colItems
		 If $objItem.NetEnabled = True Then
			GUICtrlSetData($Combo_internetSelect , "(" & $objItem.NetConnectionID & ") " & $objItem.Name & " --- " & $objItem.GUID)
		 EndIf
	  Next
	Else
	  Msgbox(0,"WMI Output","No WMI Objects Found for class: " & "Win32_NetworkAdapter" )
	Endif

EndFunc

Func runCMD($address, $x)

	;~ radio first = for overwriting on X
	If (GUICtrlRead($Radio_x86) = 1 Or $x == "86") Then
		Run("cmd /c ForceBindIP -i " & GUICtrlRead($Label_Internet) & ' ' & $address,"" , @SW_HIDE)
	ElseIf (GUICtrlRead($Radio_x64) = 1 Or $x == "64") Then
		Run("cmd /c ForceBindIP64 -i " & GUICtrlRead($Label_Internet) & ' ' & $address,"" , @SW_HIDE)
	Else
		MsgBox(0,"Select x86 | x64", "Please Select x86 or x64 or set it in .init file")
	EndIf

;~    If (GUICtrlRead($Radio_x86) = 1) Then
;~ 	  Run("cmd /c ForceBindIP -i " & GUICtrlRead($Label_Internet) & ' ' & $address,"" , @SW_HIDE)
;~    ElseIf (GUICtrlRead($Radio_x64) = 1) Then
;~ 	  Run("cmd /c ForceBindIP64 -i " & GUICtrlRead($Label_Internet) & ' ' & $address,"" , @SW_HIDE)
;~    Else
;~ 	  MsgBox(1,"Select One Option", "Please Select One Option")
;~    EndIf

EndFunc