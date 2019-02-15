
#include <GUIConstantsEx.au3>

#include "ChooseFileFolder.au3"

Global $aDrives = DriveGetDrive("ALL")
Global $sDrives = ""
For $i = 1 To $aDrives[0]
	; Only display ready drives
	If DriveStatus($aDrives[$i] & '\') <> "NOTREADY" Then $sDrives &= "|" & StringUpper($aDrives[$i])
Next
Global $sRootFolder = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", Default, -1))
ConsoleWrite($sRootFolder & @CRLF)

; Register UDF message handler
_CFF_RegMsg()

; Create GUI
$hGUI = GUICreate("_CFF_Embed Example 3 - only one TreeView active at one time", 750, 560)
GUISetBkColor(0xC4C4C4)

; Native TreeView - checkbox file only
GUICtrlCreateLabel("Press 'Start'", 10, 10, 90, 20)
$cStartButton_FileOnly = GUICtrlCreateButton("Start", 100, 10, 60, 25)
GUICtrlCreateLabel("Multiple files only", 10, 40, 200, 20)
$cTV_FileOnly = GUICtrlCreateTreeView(10, 60, 230, 400, BitOr($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
$cStop_FileOnly = GUICtrlCreateButton("Return", 10, 470, 230, 30)
$sText = "Select" & @TAB & "=  ENTER or 'Return' button" & @CRLF & "Cancel" & @TAB & "=  ESCAPE"
GUICtrlCreateLabel($sText, 10, 510, 230, 60)

; UDF TreeView - checkbox files and folders
GUICtrlCreateLabel("Press 'Start'", 260, 10, 90, 20)
$cStartButton_FilesFolders = GUICtrlCreateButton("Start", 350, 10, 60, 25)
GUICtrlCreateLabel("Multiple files and folders", 260, 40, 200, 20)
$hTV_FilesFolders = _GUICtrlTreeView_Create($hGUI, 260, 60, 230, 400, BitOr($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES))
$cStop_FilesFolders = GUICtrlCreateButton("Return", 260, 470, 230, 30)
$sText = "Select" & @TAB & "=  ENTER or 'Return' button" & @CRLF & "Cancel" & @TAB & "=  ESCAPE"
GUICtrlCreateLabel($sText, 260, 510, 230, 60)

; Native TreeView - checkbox file only with precheck
GUICtrlCreateLabel("Press 'Start'", 510, 10, 90, 20)
$cStartButton_PreCheck = GUICtrlCreateButton("Start", 600, 10, 60, 25)
GUICtrlCreateLabel("Multiple files only with preset", 510, 40, 200, 20)
$cTV_PreCheck = GUICtrlCreateTreeView(510, 60, 230, 400, BitOr($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
$cStop_PreCheck = GUICtrlCreateButton("Return", 510, 470, 230, 30)
$sText = "Select" & @TAB & "=  ENTER or 'Return' button" & @CRLF & "Cancel" & @TAB & "=  ESCAPE"
GUICtrlCreateLabel($sText, 510, 510, 230, 60)

GUISetState()

While 1

	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $cStartButton_FileOnly
			GUICtrlSetState($cStartButton_FilesFolders, $GUI_DISABLE)
			GUICtrlSetState($cStartButton_PreCheck, $GUI_DISABLE)

			; only files selectable, so only files have checkboxes
			$sSel = _CFF_Embed($cTV_FileOnly, "", "*.*", 12, $cStop_FileOnly)
			ConsoleWrite(@error & " - " & @extended & " - " & $sSel & @CRLF)

			GUICtrlSetState($cStartButton_FilesFolders, $GUI_ENABLE)
			GUICtrlSetState($cStartButton_PreCheck, $GUI_ENABLE)

		Case $cStartButton_FilesFolders
			GUICtrlSetState($cStartButton_FileOnly, $GUI_DISABLE)
			GUICtrlSetState($cStartButton_PreCheck, $GUI_DISABLE)

			; File and folders selectable so both have checkboxes
			$sSel = _CFF_Embed($hTV_FilesFolders, "", "*.*", 0 + 12 + 16, $cStop_FilesFolders)
			ConsoleWrite(@error & " - " & @extended & " - " & $sSel & @CRLF)

			GUICtrlSetState($cStartButton_FileOnly, $GUI_ENABLE)
			GUICtrlSetState($cStartButton_PreCheck, $GUI_ENABLE)

		Case $cStartButton_PreCheck
			GUICtrlSetState($cStartButton_FileOnly, $GUI_DISABLE)
			GUICtrlSetState($cStartButton_FilesFolders, $GUI_DISABLE)

			; Pre-check the AutoIt executable file
			Local $aPreCheck_List[] = [@AutoItExe]
			; Note that pre-checked file will always be returned unless expanded and cleared
			_CFF_SetPreCheck($aPreCheck_List)

			; Preset used so all checkboxes present, but only files can be selected
			; Add + 16 to $iDisplay to allow folders to be selected as well as files
			$sSel = _CFF_Embed($cTV_PreCheck, "", "*.*", 12 + 512, $cStop_PreCheck)
			ConsoleWrite(@error & " - " & @extended & " - " & $sSel & @CRLF)

			GUICtrlSetState($cStartButton_FileOnly, $GUI_ENABLE)
			GUICtrlSetState($cStartButton_FilesFolders, $GUI_ENABLE)
	EndSwitch

WEnd