
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
$hGUI = GUICreate("_CFF_Embed Example 2 - only one TreeView active at one time", 750, 560)
GUISetBkColor(0xC4C4C4)

; Native TreeView - single selection
GUICtrlCreateLabel("Choose Drive", 10, 10, 90, 20)
$cCombo_Single = GUICtrlCreateCombo("", 100, 10, 60, 20)
GUICtrlSetData(-1, $sDrives)
GUICtrlCreateLabel("Single selection", 10, 40, 200, 20)
$cTV_Single = GUICtrlCreateTreeView(10, 60, 230, 440)
$sText = "Select" & @TAB & "=  DblClk or SingleClk && ENTER" & @CRLF & "Cancel" & @TAB & "=  ESCAPE"
GUICtrlCreateLabel($sText, 10, 510, 230, 60)

; native TreeView - multiple selection (duplicates allowed) with list control, but no return control
GUICtrlCreateLabel("Choose Drive", 260, 10, 90, 20)
$cCombo_Multi_List = GUICtrlCreateCombo("", 350, 10, 60, 20)
GUICtrlSetData(-1, $sDrives)
GUICtrlCreateLabel("Multiple (duplicate) selection with list", 260, 40, 200, 20)
$cTV_Multi_List = GUICtrlCreateTreeView(260, 60, 230, 330)
$cList_Multi_List = GUICtrlCreateList("", 260, 400, 230, 100)
$sText = "Select" & @TAB & "=  DblClk" & @TAB & "Delete" & @TAB & "= Ctrl - DblClk" & @CRLF & "Return" & @TAB & "=  ENTER" & @CRLF & "Cancel" & @TAB & "=  ESCAPE"
GUICtrlCreateLabel($sText, 260, 510, 230, 60)

; UDF TreeView - multiple selection with no list control
GUICtrlCreateLabel("Press 'Start'", 510, 10, 90, 20)
$cStartButton_Multi_NoList = GUICtrlCreateButton("Start", 600, 10, 60, 25)
GUICtrlCreateLabel("Multiple selection with 'Return' button", 510, 40, 200, 20)
$hTV_Multi_NoList = _GUICtrlTreeView_Create($hGUI, 510, 60, 230, 400, BitOr($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS))
$cStop_Multi_NoList = GUICtrlCreateButton("Return", 510, 470, 230, 30)
$sText = "Select" & @TAB & "=  DblClk" & @TAB & "Delete" & @TAB & "= Ctrl - DblClk" & @CRLF & "Return" & @TAB & "=  ENTER or 'Return' button" & @CRLF & "Cancel" & @TAB & "=  ESCAPE"
GUICtrlCreateLabel($sText, 510, 510, 230, 60)

GUISetState()

While 1

	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $cCombo_Single
			; Disable combos - function is blocking so only one TreeView at a time can be used
			GUICtrlSetState($cCombo_Single, $GUI_DISABLE)
			GUICtrlSetState($cCombo_Multi_List, $GUI_DISABLE)
			GUICtrlSetState($cStartButton_Multi_NoList, $GUI_DISABLE)

			$sRoot = GUICtrlRead($cCombo_Single) & "\"

			; No return Control specified, so only single selection
			$sSel = _CFF_Embed($cTV_Single, $sRoot, "*.*", 12)
			ConsoleWrite(@error & " - " & @extended & " - " & $sSel & @CRLF)

			; Re-enable combos
			GUICtrlSetData($cCombo_Single, $sDrives)
			GUICtrlSetState($cCombo_Single, $GUI_ENABLE)
			GUICtrlSetState($cCombo_Multi_List, $GUI_ENABLE)
			GUICtrlSetState($cStartButton_Multi_NoList, $GUI_ENABLE)

		Case $cCombo_Multi_List
			GUICtrlSetState($cCombo_Single, $GUI_DISABLE)
			GUICtrlSetState($cCombo_Multi_List, $GUI_DISABLE)
			GUICtrlSetState($cStartButton_Multi_NoList, $GUI_DISABLE)

			$sRoot = GUICtrlRead($cCombo_Multi_List) & "\"

			; basic display 0 so only files selectable
			; $iDisplay + 32 so duplicate selections allowed
			; $iReturn <> 0 so multiple selection list
			; List control specified so selections displayed
			$sSel = _CFF_Embed($cTV_Multi_List, $sRoot, "*.*", 0 + 12 + 32, 1, $cList_Multi_List)
			ConsoleWrite(@error & " - " & @extended & " - " & $sSel & @CRLF)

			GUICtrlSetData($cCombo_Multi_List, $sDrives)
			GUICtrlSetState($cCombo_Single, $GUI_ENABLE)
			GUICtrlSetState($cCombo_Multi_List, $GUI_ENABLE)
			GUICtrlSetState($cStartButton_Multi_NoList, $GUI_ENABLE)

		Case $cStartButton_Multi_NoList
			GUICtrlSetState($cCombo_Single, $GUI_DISABLE)
			GUICtrlSetState($cCombo_Multi_List, $GUI_DISABLE)

			; Return control specified so multiple selections
			; Display + 16 means both files and folders can be selected
			$sSel = _CFF_Embed($hTV_Multi_NoList, "|" & $sRootFolder, "*.*", 0 + 12 + 16, $cStop_Multi_NoList)
			ConsoleWrite(@error & " - " & @extended & " - " & $sSel & @CRLF)

			GUICtrlSetState($cCombo_Single, $GUI_ENABLE)
			GUICtrlSetState($cCombo_Multi_List, $GUI_ENABLE)

	EndSwitch

WEnd