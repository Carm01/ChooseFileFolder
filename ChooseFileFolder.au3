
#include-once

; #INDEX# ============================================================================================================
; Title .........: ChooseFileFolder
; AutoIt Version : 3.3.10 +
; Language ......: English
; Description ...: Allows selection of single or multiple files and/or folders
; Remarks .......: If the script already has WM_NOTIFY or WM_COMMAND handlers then call the relevant _CFF_WM_####_Handler
;                    function from within them
; Author ........: Melba23
; Modified ......; Thanks to guinness for help with the #*@#*%# Struct !!!!
; ====================================================================================================================

;#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7

; #INCLUDES# =========================================================================================================
#include <GuiTreeView.au3>
#include <File.au3>
#include <WinAPI.au3>
; ===============================================================================================================================

; #GLOBAL VARIABLES# =================================================================================================
Global  $g_cCFF_Expand_Dummy, _			; Branch expansion event
		$g_cCFF_Select_Dummy, _			; Item selection event (checkbox on native treeView)
		$g_cCFF_Click_Dummy, _			; Item click event (checkbox on UDF TreeView)
		$g_cCFF_DblClk_Dummy, _ 		; Expansion or selection
		$g_hCFF_TreeView, _				; Active TreeView handle
		$g_hCFF_List = 9999, _			; Active list handle - with placeholder
		$g_bCFF_ActiveTV = True, _		; Flag to determine whether TreeView or list is active
		$g_bCFF_AutoExpand, _			; Flag for autoexpansion
		$g_aCFF_PreCheck = "", _		; Prechecked item list
		$g_aCFF_PreCheckRetain = ""		; Copy of precheck list if not volatile

; #CURRENT# ==========================================================================================================
; _CFF_Choose:             Creates a dialog to chose single or multiple files or folders
; _CFF_Embed:              Creates a folder tree within an existing treeview and optional list
; _CFF_SetPreCheck:        Sets a list of files/folders which will be checked when expanded
; _CFF_RegMsg:             Registers WM_NOTIFY for TreeView doubleclick and item expansion and WM_COMMAND for list focus
; _CFF_WM_NOTIFY_Handler:  WM_NOTIFY handler - reacts to doubleclick, checkbox selection and item expansion on TreeView
; _CFF_WM_COMMAND_Handler: WM_COMMAND handler - reacts list getting focus
; ====================================================================================================================

; #INTERNAL_USE_ONLY#=================================================================================================
; _CFF_Fill_Combo:      Creates and fills a combo for drive selection
; _CFF_Fill_Drives:     Fills a TreeView with ready drives
; _CFF_Fill_Branch:     Fills a TreeView branch with folders and files on expansion
; _CFF_AutoExpand:      Expand tree to defined folder on start
; _CFF_File_Visible:    Ensure files visible if selecting files only and displaying both files and folders
; _CFF_Check_Display:   Checks for valid Display parameter
; _CFF_Check_Selection: Checks selection is valid
; _CFF_ParseTV:         Creates array of current TreeView checkbox state
; _CFF_Adjust_Parents:  Adjusts parent checkboxes
; _CFF_Adjust_Children: Adjust children to parent checked state
; _CFF_List_Add:        Adds item to return list
; _CFF_List_Del:        Deletes item from return list
; ====================================================================================================================

; #FUNCTION# =========================================================================================================
; Name...........: _CFF_Choose
; Description ...: Creates a dialog to chose single or multiple files or folders
; Syntax.........: _CFF_Choose ($sTitle, [$iW = 1, [$iH = 1, [$iX = -1, [$iY = -1, [$sRoot = "", [$s_Mask = "", [$iDisplay = 0, [$iMultiple = True, [$hParent = 0]]]]]]]]])
; Parameters ....: $sTitle      - Title of dialog - only necessary parameter!
;                  $iW, $iH     - Width, Height parameters for dialog (Default and minimum = 250 x 300).  Set negative for resizable dialog
;                  $iX, $iY     - Left, Top parameters for dialog (Default = centred)
;                  $sRoot       - Valid path = Tree to display
;                                   If folder, contents of folder are displayed
;                                   If filename, contents of parent folder are displayed with the file highlighted
;                                 To permit drive selection:
;                                   "" (default) = All ready drives shown
;                                   List of drive letters (e.g. "cde") = Limit display to these drives if ready
;                                     Add "|drive letter" (e.g. "cde|d" or "|d") = Highlight and expand specified drive tree on opening
;                                      or "|folder path" (e.g. "cde|d:\folder) = Highlight specified folder (add "\" to expand folder)
;                                     Add "|c" (e.g. cde||c" or "|d|c" or "||c)  = Display drives in separate combo not treeview
;                  $sMask       - Filter for result. Multiple filters must be separated by ";"
;                                 Use "|" to separate 3 possible sets of filters: "Include|Exclude|Exclude_Folders"
;                                   Include = Files/Folders to include (default = "*" [all])
;                                   Exclude = Files/Folders to exclude (default = "" [none])
;                                   Exclude_Folders = only used if the basic $iDisplay = 0 to exclude defined folders (default = "" [none])
;                  $iDisplay    - Determine what is displayed in the tree
;                                 0     - Entire folder tree and all matching files within folders - only files can be selected
;                                 1     - All matching files within the specified folder - subfolders are not displayed
;                                 2     - Entire folder tree only - no files.  Doubleclicks will expand, not select, item
;                                 3     - As 2 but doubleclicks will select item if it has no children
;                                 + 4   - Do not display Hidden files/folders
;                                 + 8   - Do not display System files/folders
;                                 + 16  - Both files and folders selectable if basic $iDisplay set to 0 - folders have trailing \
;                                 + 32  - Allow duplicate selections when clicking
;                                 + 64  - Hide file extensions (only valid if file mask specifies a single extension)
;                                 + 128 - Scroll to first file in folder - only valid if basic $iDisplay set to 0
;                                 + 256 - Display splashscreens when searching for drives and autoexpanding tree
;                                 + 512 - When using checkboxes only return deepest items on path - default return all checked items
;                  $iMultiple   - 0 or non-numeric = Only 1 selection (default)
;                                 -1               = Multiple selections using checkboxes
;                                 10-50            = Multiple selections added to list.  List size 10-50% of dialog
;                                 Other numeric    = Multiple selections added to list.  List size set to 20% of dialog
;                  $hParent     - Handle of GUI calling the dialog, (default = 0 - no parent GUI)
; Requirement(s).: v3.3 +
; Return values .: Success: String containing selected items - multiple items delimited by "|"
;                  Failure: Returns "" and sets @error as follows:
;                           1 = Invalid $sRoot parameter
;                                       @extended = 1 - Path does not exist
;                                                   2 - Invalid drive list
;                           2 = Invalid $iDisplay parameter
;                           3 = Invalid $hParent parameter
;                           4 = Dialog creation failure
;                           5 = Cancel button, {ESCAPE} or GUI [X] pressed
; Author ........: Melba23
; Modified ......:
; Remarks .......: Multiple selections:
;                      If using list:
;                          Press "Add" or doubleclick in treeview to add item to list
;                          Press "Delete" to delete selected item from treeview or list
;                      If using checkboxes:
;                          Use checkboxes to select items
;                          Pressing Ctrl when (un)checking sets all children to same state
;                          Pressing Alt when (un)checking sets parents to same state (uncheck if no other checked items on same path)
;                      Both cases:
;                          Press "Return" button or "{ENTER}" key when selection ended
;                          Press "Cancel" button or "{ESCAPE}" key to abandon selection
; Example........: Yes
;=====================================================================================================================
Func _CFF_Choose($sTitle, $iW = 1, $iH = 1, $iX = -1, $iY = -1, $sRoot = "", $sMask = "*", $iDisplay = 0, $iMultiple = 0, $hParent = 0)

	Local Const $iFixed = 0x80C80000     ; BitOR($WS_POPUPWINDOW, $WS_CAPTION)
	Local Const $iResizable = 0x80CC0000 ; BitOR($WS_POPUPWINDOW, $WS_CAPTION, $WS_SIZEBOX)
	Local Const $iComposite = 0x02000000 ; $WS_EX_COMPOSITED

	Local $bShow_Ext = True, $iHide_HS = 0, $bBoth_Selectable = False, $bDuplicates_Allowed = False, $bFileScroll = False
	Local $bCombo = False, $bDefFolder_Open = False, $bSplash = False, $bDeepest = False
	Local $sDrives = "", $sDefDrive = "", $sDefFolder_Check, $sDefFolder, $sDefFile
	Local $iStyle = $iFixed, $iExStyle = 0, $iRedraw_Count = 5, $bNoFolderCheck = False
	Local $aAll_Drives[1] = [0], $aNetwork_Drives[1] = [0], $aFill_Ret[3], $aTVCheckData

	; Set default autoexpand value
	$g_bCFF_AutoExpand = False

	; Check size
	If $iW = Default Then
		$iW = 1
	EndIf
	If $iH = Default Then
		$iH = 1
	EndIf

	; Set dialog resizing style
	If $iW < 0  Or $iH < 0 Then
		$iStyle = $iResizable
		$iExStyle = $iComposite
		$iW = Abs($iW)
		$iH = Abs($iH)
	EndIf

	; Check position
	If $iX = Default Then
		$iX = -1
	EndIf
	If $iY = Default Then
		$iY = -1
	EndIf

	; Check path
	Switch $sRoot
		Case "", Default
			$sRoot = ""
		Case Else
			; If a path
			If FileExists($sRoot) Then
				; Check if file or folder
				If StringInStr(FileGetAttrib($sRoot), "D") Then
					; Add trailing \ if needed
					If StringRight($sRoot, 1) <> "\" Then
						$sRoot &= "\"
					EndIf
				Else
					$sDefFile = StringRegExpReplace($sRoot, "^.*\\", "")
					$sRoot = StringRegExpReplace($sRoot, "(^.*\\)(.*)", "\1")
				EndIf
			Else
				; Split to look for eventual drive|default|combo info
				Local $aSplit = StringSplit($sRoot, "|")
				If Not @error And $aSplit[0] < 4 Then
					; Check required drive list
					If Not StringRegExp($aSplit[1], "(?i)^[a-z]*$") Then
						; Reset precheck array as required
						$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
						Return SetError(1, 2, "")
					EndIf
					; And store
					$sDrives =StringUpper($aSplit[1])
					; Check for default drive
					If $aSplit[0] > 1 Then
						$sDefDrive = StringUpper(StringLeft($aSplit[2], 1))
						; Check if default in drive list
						If $sDrives And (Not StringInStr($sDrives, $sDefDrive)) Then
							; Clear if not
							$sDefDrive = ""
						Else
							; Add colon
							If $sDefDrive Then
								$sDefDrive &= ":"
							EndIf
						EndIf
						; Is it a valid folder path?
						If FileExists($aSplit[2]) And StringInStr(FileGetAttrib($aSplit[2]), "D") Then
							$sDefFolder = $aSplit[2]
							; Check if trailing backslash
							If StringRight($sDefFolder, 1) = "\" Then
								; Set flag
								$bDefFolder_Open = True
								; Remove the trailing backslash
								$sDefFolder = StringTrimRight($sDefFolder, 1)
							EndIf
							; Replace backslashes for comparison when expanding
							$sDefFolder_Check = StringReplace($sDefFolder, "\", "|")
						EndIf
					EndIf
					; Use combo for drive choice?
					If $aSplit[0] = 3 And $aSplit[3] = "C" Then
						$bCombo = True
					EndIf
					; Clear root parameter to show drive choice required
					$sRoot = ""
				Else
					; Reset precheck array as required
					$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
					; Not valid path or drive list
					Return SetError(1, 1, "")
				EndIf
			EndIf
	EndSwitch

	; Create drive lists if no root set
	If $sRoot = "" Then
		; Check if drive list set
		If $sDrives Then
			; Create array of these drives
			$aAll_Drives = StringSplit($sDrives, "")
			For $i = 1 To $aAll_Drives[0]
				$aAll_Drives[$i] &= ":"
			Next
		Else
			; Get array of drives
			$aAll_Drives = DriveGetDrive("ALL")
			If @error Then
				Local $aAll_Drives[1] = [0]
			EndIf
		EndIf
		; Get array of network drives - to show indexing label if required
		$aNetwork_Drives = DriveGetDrive("NETWORK")
		If @error Then
			Local $aNetwork_Drives[1] = [0]
		EndIf
	EndIf

	; Check Mask
	If $sMask = Default Or $sMask = "" Then
		$sMask = "*"
	EndIf

	; Extract File mask (needed for hide extension code)
	Local $aMaskSplit = StringSplit($sMask, "|")
	Local $sFile_Mask = $aMaskSplit[1]

	; Check Display parameter
	If $iDisplay = Default Then
		$iDisplay = 0
	EndIf
	$iDisplay = _CFF_Check_Display($iDisplay, $sFile_Mask, $bShow_Ext, $iHide_HS, $bBoth_Selectable, $bDuplicates_Allowed, $bFileScroll, $bSplash, $bDeepest)
	If @error Then
		; Reset precheck array as required
		$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
		Return SetError(2, 0, "")
	EndIf

	; Verify precheck list if set
	If IsArray($g_aCFF_PreCheck) Then
		_CFF_Verify_PreCheck($g_aCFF_PreCheck, $iDisplay, $bBoth_Selectable)
	EndIf

	; Check selection type and list size
	Local $bSingle_Sel = False, $bCheckBox = False, $iList_Percent = .2
	If $iMultiple = Default Then $iMultiple = 0
	Switch Number($iMultiple)
		Case 0, Default
			$bSingle_Sel = True
		Case 10 To 50
			$iList_Percent = .01 * $iMultiple
		Case -1
			$bCheckBox = True
			If $iDisplay = 3 Then $iDisplay = 2 ; Remove possible "selection on doubleclick"
	EndSwitch

	; Check if checkboxes required
	If $iMultiple = -1 Then
		; Check if folder checkboxes should be hidden - only files selectable and no prechecks
		If $iDisplay = 0 And Not($bBoth_Selectable) And Not(IsArray($g_aCFF_PreCheck)) Then
			$bNoFolderCheck = True
		EndIf
	EndIf

	; Check parent
	Switch $hParent
		Case Default
			$hParent = 0
		Case 0
		Case Else
			If Not IsHWnd($hParent) Then
				; Reset precheck array as required
				$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
				Return SetError(3, 0, "")
			EndIf
	EndSwitch

	; Check for width and height minima
	If $bSingle_Sel Then
		If $iW < 250 Then $iW = 250
	Else
		If $iW < 350 Then $iW = 350
	EndIf
	If $iH < 300 Then $iH = 300
	; Set button size
	Local $iButton_Width = Int(($iW - 50) / 4)
	If $iButton_Width > 80 Then $iButton_Width = 80
	If $iButton_Width < 50 Then $iButton_Width = 50

	; Create Dialog
	Local $hCFF_Win = GUICreate($sTitle, $iW, $iH, $iX, $iY, $iStyle, $iExStyle, $hParent)
	If @error Then
		; Reset precheck array as required
		$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
		Return SetError(4, 0, "")
	EndIf
	GUISetBkColor(0xCECECE)

	If $bSplash Then
		; Show splash if dialog display delay likely - multiple drives or autoexpansion
		Local $sMsg = ""
		If Not $sRoot And Not $sDefFolder Then
			$sMsg = "Detecting drives..."
		ElseIf $sDefFolder Then
			$sMsg = "Expanding tree..."
			; Prevent user input
			BlockInput(1)
		EndIf
		If $sMsg Then
			Local $aPos = WinGetPos($hCFF_Win)
			SplashTextOn("ChooseFileFolder", $sMsg, $iW - 100, 100, $aPos[0] + 50, $aPos[1] + 100)
		EndIf
	EndIf

	; Declare variables
	Local $cCan_Button, $cSel_Button = 9999, $cAdd_Button = 9999, $cDel_Button = 9999, $cRet_Button = 9999
	Local $cTreeView = 9999, $cList = 9999, $cDrive_Combo = 9999
	Local $sCurrDrive = "", $sSelectedPath, $iList_Height, $iTV_Height, $sAddFile_List = ""
	Local $cItem, $hItem

	; Create buttons
	If $bSingle_Sel Then
		$cSel_Button = GUICtrlCreateButton("Select", $iW - ($iButton_Width + 10), $iH - 40, $iButton_Width, 30)
	Else
		If Not $bCheckBox Then
			$cAdd_Button = GUICtrlCreateButton("Add", 10, $iH - 40, $iButton_Width, 30)
			$cDel_Button = GUICtrlCreateButton("Delete", $iButton_Width + 20, $iH - 40, $iButton_Width, 30)
		EndIf
		$cRet_Button = GUICtrlCreateButton("Return", $iW - ($iButton_Width + 10), $iH - 40, $iButton_Width, 30)
	EndIf
	$cCan_Button = GUICtrlCreateButton("Cancel", $iW - ($iButton_Width + 10) * 2, $iH - 40, $iButton_Width, 30)
	Local $cRet_Dummy = GUICtrlCreateDummy() ; fires on ENTER
	Local $cEsc_Dummy = GUICtrlCreateDummy() ; fires on ESC
	; Set accel keys
	Local $aAccelKeys[2][2] = [["{ENTER}", $cRet_Dummy],["{ESC}", $cEsc_Dummy]]
	GUISetAccelerators($aAccelKeys)

	; Create controls
	Local $iOffset = 30 ; Offset if combo used
	If $bSingle_Sel Then
		If $bCombo Then ; Combo and TV
			; Create and fill Combo
			$cDrive_Combo = _CFF_Fill_Combo($iW, $sDrives, $sDefDrive)
		Else ; TV only
			; No offset
			$iOffset = 0
		EndIf
		; Create TV
		$cTreeView = GUICtrlCreateTreeView(10, 10 + $iOffset, $iW - 20, $iH - 60 - $iOffset)
	Else
		If $bCombo Then ; Combo and TV
			; Create and fill Combo
			$cDrive_Combo = _CFF_Fill_Combo($iW, $sDrives, $sDefDrive)
		Else ; TV only
			; No offset
			$iOffset = 0
		EndIf
		If $bCheckBox Then
			; Create TV
			$cTreeView = GUICtrlCreateTreeView(10, 10 + $iOffset, $iW - 20, $iH - 60 - $iOffset, BitOr($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
		Else
			; Calculate control heights
			$iList_Height = Int(($iH - 60 - $iOffset) * $iList_Percent)
			If $iList_Height < 40 Then
				$iList_Height = 40
			EndIf
			$iTV_Height = $iH - $iList_Height - 60 - $iOffset
			; Create TV
			$cTreeView = GUICtrlCreateTreeView(10, 10 + $iOffset, $iW - 20, $iTV_Height)
			; Create List
			$cList = GUICtrlCreateList("", 10, 10 + $iOffset + $iTV_Height, $iW - 20, $iList_Height, 0x00A00101) ; BitOR($WS_BORDER, $WS_VSCROLL, $LBS_NOINTEGRALHEIGHT, $LBS_NOTIFY)
		EndIf
	EndIf

	; Create dummy control to fire when [+] clicked
	$g_cCFF_Expand_Dummy = GUICtrlCreateDummy()
	; Create dummy control to fire when item checked
	$g_cCFF_Select_Dummy = GUICtrlCreateDummy()
	; Create dummy control to fire on DblClk
	$g_cCFF_DblClk_Dummy = GUICtrlCreateDummy()
	; Create a dummy control to redraw TreeViw if auto-expand fails
	Local $cRedraw_Dummy = GUICtrlCreateDummy()

	; Set Global values for handler checks
	$g_hCFF_TreeView = GUICtrlGetHandle($cTreeView)
	$g_hCFF_List = GUICtrlGetHandle($cList)

	; Set resizing if required
	If $iStyle = $iResizable Then
		GUICtrlSetResizing($cTreeView, 2 + 4 + 32)  ; $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKRIGHT
		GUICtrlSetResizing($cList, 2 + 4)           ; $GUI_DOCKLEFT + $GUI_DOCKRIGHT
		GUICtrlSetResizing($cAdd_Button, 2 + 64)    ; $GUI_DOCKLEFT + $GUI+DOCKBOTTOM
		GUICtrlSetResizing($cDel_Button, 8 + 64)    ; $GUI_DOCKHCENTER + $GUI+DOCKBOTTOM
		GUICtrlSetResizing($cCan_Button, 8 + 64)    ; $GUI_DOCKHCENTER + $GUI+DOCKBOTTOM
		GUICtrlSetResizing($cSel_Button, 4 + 64)    ; $GUI_DOCKRIGHT + $GUI+DOCKBOTTOM
		GUICtrlSetResizing($cRet_Button, 4 + 64)    ; $GUI_DOCKRIGHT + $GUI+DOCKBOTTOM
	EndIf

	; Hide dialog until tree drawing complete
	WinSetTrans($hCFF_Win, "", 1) ; Must be 1 as 0 = hidden
	; Display dialog so that autoexpansion can occur if required
	GUISetState()
	; Force OnTop - near invisible can be hidden
	WinSetOnTop($hCFF_Win, "", 1)

	; Set required default drive if required
	If $sDefDrive Then
		If $bCombo Then
			; Set default drive as root
			$sRoot = $sDefDrive & "\"
			; Show in combo
			GUICtrlSetData($cDrive_Combo, $sDefDrive)
		EndIf
	EndIf

	; Fill tree
	If $sRoot Then
		; If root folder specified then fill TV
		$aFill_Ret = _CFF_Fill_Branch($cTreeView, $cTreeView, $sRoot, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)
	Else
		If $bCombo Then
			_CFF_Fill_Combo($iW, $sDrives, $sDefDrive)
		Else
			_CFF_Fill_Drives($cTreeView, $aAll_Drives, $sDefDrive, $sDefFolder, $bDeepest, $bNoFolderCheck)
		EndIf
	EndIf

	; Scroll to show files if required
	If $sDefFile Then
		; Ensure default file visible, highlighted and selected
		$hItem = _GUICtrlTreeView_GetFirstItem($cTreeView)
		Do
			If _GUICtrlTreeView_GetText($cTreeView, $hItem) = $sDefFile Then
				_GUICtrlTreeView_EnsureVisible($cTreeView, $hItem)
				_GUICtrlTreeView_ClickItem($cTreeView, $hItem)
				ExitLoop
			EndIf
			$hItem = _GUICtrlTreeView_GetNextChild($cTreeView, $hItem)
		Until $hItem = 0
	Else
		; If file scroll selected and files found
		If $bFileScroll And $aFill_Ret[2] Then
			; Scroll to first file in this branch
			_CFF_File_Visible($cTreeView, $aFill_Ret)
		EndIf
	EndIf

	; If defined folder
	If $sDefFolder Then
		; Ensure auto expansion
		GUICtrlSendToDummy($cRedraw_Dummy)
		; Set counter
		$iRedraw_Count = 0
	EndIf

	; Parse TV for check data
	If $bCheckBox Then
		$aTVCheckData = _CFF_ParseTV($cTreeView)
	EndIf

	; Change to MessageLoop mode
	Local $nOldOpt = Opt('GUIOnEventMode', 0)

	While 1

		; Run when any expansion ended or by default
		If $iRedraw_Count = 5 Then
			; Reshow dialog
			WinSetTrans($hCFF_Win, "", 255)
			; Cancel OnTop
			WinSetOnTop($hCFF_Win, "", 0)
			; Reset flag
			$iRedraw_Count = 0
			; Delete Splash
			SplashOff()
			; Restore user input
			BlockInput(0)
		EndIf

		Switch GUIGetMsg()
			Case $cCan_Button, $cEsc_Dummy, -3 ; $GUI_EVENT_CLOSE
				GUIDelete($hCFF_Win)
				; Restore previous mode
				Opt('GUIOnEventMode', $nOldOpt)
				; Reset precheck array as required
				$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
				; And exit
				Return SetError(5, 0, "")

			Case $cRedraw_Dummy
				; Empty TreeView
				_GUICtrlTreeView_DeleteAll($cTreeView)
				; Refill drives
				_CFF_Fill_Drives($cTreeView, $aAll_Drives, $sDefDrive, $sDefFolder, $bDeepest, $bNoFolderCheck)
				; Increase counter
				$iRedraw_Count += 1
				If $iRedraw_Count < 5 Then
					; Re-expand to default folder
					_CFF_AutoExpand($cTreeView, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $sDefFolder, $bDefFolder_Open, $iRedraw_Count, $bDeepest, $bNoFolderCheck, $bCombo, True) ; Native TreeView
					; Check expansion correct
					If _GUICtrlTreeView_GetTree($cTreeView) <> $sDefFolder_Check Then
						; Force another redraw
						GUICtrlSendToDummy($cRedraw_Dummy)
					Else
						$iRedraw_Count = 5
					EndIf
				EndIf
				If $bCheckBox Then
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($cTreeView)
				EndIf

			Case $g_cCFF_DblClk_Dummy
				If $bCheckBox Then
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($cTreeView)
				Else
					; Check doubleclick selection is permitted (not default for folders only)
					If $iDisplay = 2 Then
						; Not permitted
					ElseIf $iDisplay = 3 Then
						; Check if any children
						$cItem = GUICtrlRead($cTreeView)
						$hItem = GUICtrlGetHandle($cItem)
						If _GUICtrlTreeView_GetChildCount($g_hCFF_TreeView, $hItem) < 1 Then
							; No children so selection permitted
							ContinueCase
						EndIf
					Else
						; Permitted
						ContinueCase
					EndIf
				EndIf

			Case $cSel_Button, $cAdd_Button
				; Get item data
				$cItem = GUICtrlRead($cTreeView)
				$hItem = GUICtrlGetHandle($cItem)
				; Check path is a valid selection
				$sSelectedPath = _CFF_Check_Selection($g_hCFF_TreeView, $hItem, $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
				If $sSelectedPath Then
					; Valid item selected
					If $bSingle_Sel Then
						GUIDelete($hCFF_Win)
						; Restore previous mode
						Opt('GUIOnEventMode', $nOldOpt)
						; Reset precheck array as required
						$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
						; Return valid path
						Return $sSelectedPath
					Else
						$sAddFile_List = _CFF_List_Add($sAddFile_List, $sSelectedPath, $cList, $bDuplicates_Allowed, $cTreeView)
					EndIf
				EndIf

			Case $cDel_Button
				; Check if treeView is active
				If $g_bCFF_ActiveTV Then
					; Get TreeView item selected
					$cItem = GUICtrlRead($cTreeView)
					$hItem = GUICtrlGetHandle($cItem)
					; Check path is a valid selection
					$sSelectedPath = _CFF_Check_Selection($g_hCFF_TreeView, $hItem, $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
					If $sSelectedPath Then
						; Valid item selected
						$sAddFile_List = _CFF_List_Del($sAddFile_List, $sSelectedPath, $cList, $cTreeView)
					EndIf
				Else
					; Delete list item selected
					$sAddFile_List = _CFF_List_Del($sAddFile_List, GUICtrlRead($cList), $cList, $cTreeView)
					; Set flag
					$g_bCFF_ActiveTV = True
				EndIf

			Case $g_cCFF_Select_Dummy
				If $bCheckBox Then
					Local $bState, $iItemIndex
					; Get handle of selected item
					$hItem = GUICtrlGetHandle(GUICtrlRead($cTreeView))

					; Determine If item path is a valid selection
					$sSelectedPath = _CFF_Check_Selection($cTreeView, $hItem, $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
					; If not valid selection
					If Not $sSelectedPath Then
						; If prechecks exist
						If IsArray($g_aCFF_PreCheck) Then
							; Get stored item check state
							$iItemIndex = _ArraySearch($aTVCheckData, $hItem)
							; And force to that state
							_GUICtrlTreeView_SetChecked($cTreeView, $hItem, $aTVCheckData[$iItemIndex][1])
						Else
							; If folder checkboxes are shown
							If Not($bNoFolderCheck) Then
								; Force to unchecked
								_GUICtrlTreeView_SetChecked($cTreeView, $hItem, False)
							EndIf
						EndIf
					EndIf

					; Check if Control or Alt pressed
					_WinAPI_GetAsyncKeyState(0x11) ; Needed to avoid double setting
					_WinAPI_GetAsyncKeyState(0x12)
					If _WinAPI_GetAsyncKeyState(0x11) Or _WinAPI_GetAsyncKeyState(0x12) Then
						; Determine checked state
						$bState = _GUICtrlTreeView_GetChecked($cTreeView, $hItem)
						; Find item in array
						$iItemIndex = _ArraySearch($aTVCheckData, $hItem)
						; If checked state has altered
						If $aTVCheckData[$iItemIndex][1] <> $bState Then
							; Store new state
							$aTVCheckData[$iItemIndex][1] = $bState
							; Adjust parents if Alt pressed and only deepest item is to be returned
							If $bDeepest And _WinAPI_GetAsyncKeyState(0x12) Then
								_CFF_Adjust_Parents($g_hCFF_TreeView, $hItem, $aTVCheckData, $bState)
							EndIf
							; Adjust visible children if Ctrl pressed
							If _WinAPI_GetAsyncKeyState(0x11) Then
								_CFF_Adjust_Children($cTreeView, $hItem, $aTVCheckData, $bState)
								; Expand item if required
								If Not _GUICtrlTreeView_GetExpanded($cTreeView, $hItem) Then
									GUICtrlSendToDummy($g_cCFF_Expand_Dummy, $hItem)
								EndIf
							EndIf
						EndIf
					EndIf
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
				EndIf

			Case $cRet_Button, $cRet_Dummy
				If $bCheckBox Then
					; Construct return list
					Local $sOldSep = Opt("GUIDataSeparatorChar", "\"), $sFullPath
					$sAddFile_List = ""
					For $i = 0 To UBound($aTVCheckData) - 1
						If $aTVCheckData[$i][1] And $aTVCheckData[$i][2] Then
							; Check if valid selection
							$sFullPath = _CFF_Check_Selection($g_hCFF_TreeView, $aTVCheckData[$i][0], $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
							If $sFullPath Then $sAddFile_List &= $sFullPath & "|"
						EndIf
					Next
					Opt("GUIDataSeparatorChar", $sOldSep)

					; Check if only deepest items are to be returned
					If $bDeepest Then
						; Convert list to array
						Local $aDeepest = StringSplit($sAddFile_List, "|")
						; Loop through array checking if item forms part of preceding item
						For $i = $aDeepest[0] - 2 To 1 Step -1
							$sFullPath = $aDeepest[$i]
							If StringInStr(FileGetAttrib($sFullPath), "D") Then
								If StringInStr($aDeepest[$i + 1], $aDeepest[$i]) Then
									; Remove item from list
									$sAddFile_List = StringReplace($sAddFile_List, $aDeepest[$i] & "|", "")
								EndIf
							EndIf
						Next
						; If there were prechecked items
						If IsArray($g_aCFF_PreCheck) Then
							; Loop through to remove any parent items that were not fully expanded
							For $i = 1 To $aDeepest[0]
								For $j = 0 To UBound($g_aCFF_PreCheck) - 1
									If StringInStr($g_aCFF_PreCheck[$j], $aDeepest[$i]) Then
										$sAddFile_List = StringReplace($sAddFile_List, $aDeepest[$i] & "|", "")
									EndIf
								Next
							Next
						EndIf
					EndIf
					; Now add unused prechecked items
					For $i = 0 To UBound($g_aCFF_PreCheck) - 1
						If $g_aCFF_PreCheck[$i] Then
							$sAddFile_List &= $g_aCFF_PreCheck[$i] & "|"
						EndIf
					Next
				EndIf
				; Close dialog
				GUIDelete($hCFF_Win)
				; Restore previous mode
				Opt('GUIOnEventMode', $nOldOpt)
				; Reset precheck array as required
				$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
				; Remove final | from return string and return
				Return StringTrimRight($sAddFile_List, 1)

			Case $g_cCFF_Expand_Dummy
				; Get expanded item hamdle
				$hItem = GUICtrlRead($g_cCFF_Expand_Dummy)
				If $hItem Then
					; Select item
					_GUICtrlTreeView_ClickItem($g_hCFF_TreeView, $hItem)
					; Get expanded item data
					$cItem = GUICtrlRead($cTreeView)
					$hItem = GUICtrlGetHandle($cItem)
					$sSelectedPath = $sRoot & StringReplace(_GUICtrlTreeView_GetTree($cTreeView, $hItem), "|", "\")
					; Check if dummy child exists or has already been filled
					Local $hFirstChild = _GUICtrlTreeView_GetFirstChild($g_hCFF_TreeView, $hItem)
					Local $sFirstChild = _GUICtrlTreeView_GetText($g_hCFF_TreeView, $hFirstChild)
					; If dummy child exists
					If $sFirstChild = "" Then
						; Fill with content
						$aFill_Ret = _CFF_Fill_Branch($cTreeView, $cItem, $sSelectedPath, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)
						; Delete the dummy
						_GUICtrlTreeView_Delete($g_hCFF_TreeView, $hFirstChild)
					EndIf
					; If file scroll selected AND files were found AND the branch is being expanded
					If $bFileScroll And $aFill_Ret[2] And _GUICtrlTreeView_GetExpanded($g_hCFF_TreeView, $hItem) Then
						; Scroll to first file in this branch
						_CFF_File_Visible($g_hCFF_TreeView, $aFill_Ret)
					EndIf
					; Clear the flag to reactivate the handler
					GUICtrlSendToDummy($g_cCFF_Expand_Dummy, 0)
				EndIf
				If $bCheckBox Then
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
				EndIf

			Case $cDrive_Combo
				If GUICtrlRead($cDrive_Combo) <> $sCurrDrive Then
					; Get drive chosen
					$sCurrDrive = GUICtrlRead($cDrive_Combo)
					If $sRoot Then
						; Delete current content
						_GUICtrlTreeView_DeleteAll($g_hCFF_TreeView)
					Else
						; Show TV
						GUICtrlSetState($cTreeView, 16) ; $GUI_SHOW
					EndIf
					; Set root path
					$sRoot = $sCurrDrive & "\"
					; Fill TV
					$aFill_Ret = _CFF_Fill_Branch($cTreeView, $cTreeView, $sRoot, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)
					; If file scroll selected and files found
					If $bFileScroll And $aFill_Ret[2] Then
						; Scroll to first file in this branch
						_CFF_File_Visible($cTreeView, $aFill_Ret)
					EndIf
				EndIf
		EndSwitch

	WEnd

EndFunc   ;==>_CFF_Choose

; #FUNCTION# =========================================================================================================
; Name...........: _CFF_Embed
; Description ...: Creates the folder tree within an existing treeview and optional list
; Syntax.........: _CFF_Embed ($cTreeView, [$sRoot = ""[, $sMask = "*"[, $iDisplay = 0[, $cReturn = 0[, $cList = 0]]]])
; Parameters ....: $cTreeView   - Handle/ControlID of TreeView to use - only required parameter!
;                  $sRoot       - Path tree to display
;                                   If a filename is passed, the tree for the folder opens with the file highlighted
;                                 To permit drive selection:
;                                   "" (default) = All ready drives shown
;                                   List of drive letters (e.g. "cde") = Limit display to these drives if ready
;                                     Add "|drive letter" (e.g. "cde|d" or "|d") = Highlight and expand specified drive tree on opening
;                                      or "|folder path" (e.g. "cde|d:\folder) = Highlight specified folder (add "\" to expand folder)
;                  $sMask       - Filter for result. Multiple filters must be separated by ";"
;                                 Use "|" to separate 3 possible sets of filters: "Include|Exclude|Exclude_Folders"
;                                   Include = Files/Folders to include (default = "*" [all])
;                                   Exclude = Files/Folders to exclude (default = "" [none])
;                                   Exclude_Folders = only used if the basic $iDisplay = 0 to exclude defined folders (default = "" [none])
;                  $iDisplay    - Determine what is displayed in the tree
;                                 0     - Entire folder tree and all matching files within folders - only files can be selected (default)
;                                 1     - All matching files within the specified folder - subfolders are not displayed
;                                 2     - Entire folder tree only - no files.  Doubleclicks will expand, not select, item
;                                 3     - As 2 but doubleclicks will select item if it has no children
;                                 + 4   - Do not display Hidden files/folders
;                                 + 8   - Do not display System files/folders
;                                 + 16  - Both files and folders can be selected - only valid if $iDisplay set to 0
;                                 + 32  - Allow duplicate selections when clicking
;                                 + 64  - Hide file extensions (only valid if file mask specifies a single extension)
;                                 + 128 - Scroll to first file in folder - only valid if basic $iDisplay set to 0
;                                 + 512 - When using checkboxes only return deepest items on path - default return all checked items
;                  $cReturn     - Set single/multiple selection
;                                 0 -  Single selection (default).  Doubleclick returns item - @extended set to 0
;                                 1  - Multiple selection.  Doubleclick adds item to return string
;                                 Valid ControlID - Control to action to end multiple selection - @extended set to 0
;                  $cList       - [Optional] ControlID of list to fill with multiple selections
; Requirement(s).: v3.3 +
; Return values .: Success: Single selection = String containing selected item
;                           Multi selection  = String of selected items delimited by "|"
;                  Failure: Returns "" and sets @error as follows:
;                           1 = Invalid TreeView ControlID
;                           2 = Invalid $sPath parameter
;                           3 = Invalid $iDisplay parameter
;                           4 = Invalid $cList ControlID
;                           5 = Invalid or missing $cReturn ControlID
;                           6 = Empty tree
;                           7 = Esc pressed to exit
; Author ........: Melba23
; Modified ......:
; Remarks .......: - Press "ESC" to cancel selection process
;                  - For multiple selection:
;                      If list:
;                          Doubleclick treeview item to add to list
;                          Press Ctrl while doubleclicking to delete last instance of item
;                      If checkboxes:
;                          Use checkboxes to select items
;                          Pressing Ctrl when (un)checking sets all children to same state
;                          Pressing Alt when (un)checking sets parents to same state if no other checked boxes on path
;                      Use "{ENTER}" key or passed control to return selection - @extended set to 1
; Example........: Yes
;=====================================================================================================================
Func _CFF_Embed($cTreeView, $sRoot = "", $sMask = "*", $iDisplay = 0, $cReturn = 0, $cList = 0)

	Local $bSingle_Sel, $cItem, $hItem, $sSelectedPath, $sAddFile_List = ""
	Local $sDrives = "", $sDefDrive = "", $sDefFolder, $sDefFolder_Check, $sDefFile, $aFill_Ret[3], $aAll_Drives, $aNetwork_Drives
	Local $bShow_Ext = True, $iHide_HS = 0, $bBoth_Selectable = False, $bDuplicates_Allowed = False, $bFileScroll = False, $bSplash = False, $bDeepest = False
	Local $iRedraw_Count, $bDefFolder_Open = False, $bCheckBox = False, $aTVCheckData, $aOldCheckData, $bNoFolderCheck = False

	; Set default autoexpand value
	$g_bCFF_AutoExpand = False

	; Check treeview type and set TV_ID
	Local $bNative_TV = True, $vTV_ID = $cTreeView
	If IsHWnd($cTreeView) Then
		$g_hCFF_TreeView = $cTreeView
		$bNative_TV = False
		$vTV_ID = $g_hCFF_TreeView
	Else
		If Not IsHWnd(GUICtrlGetHandle($cTreeView)) Then
			; Reset precheck array as required
			$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
			Return SetError(1, 0, "")
		Else
			$g_hCFF_TreeView = GUICtrlGetHandle($cTreeView)
		EndIf
	EndIf

	; Check path
	Switch $sRoot
		Case "", Default
			$sRoot = ""
		Case Else
			; If a path
			If FileExists($sRoot) Then
				; Check if file or folder
				If StringInStr(FileGetAttrib($sRoot), "D") Then
					; Add trailing \ if needed
					If StringRight($sRoot, 1) <> "\" Then
						$sRoot &= "\"
					EndIf
				Else
					$sDefFile = StringRegExpReplace($sRoot, "^.*\\", "")
					$sRoot = StringRegExpReplace($sRoot, "(^.*\\)(.*)", "\1")
				EndIf
			Else
				; Split to look for eventual drive|default info - ignore combo section
				Local $aSplit = StringSplit($sRoot, "|")
				If Not @error And $aSplit[0] < 4 Then
					; Check required drive list
					If Not StringRegExp($aSplit[1], "(?i)^[a-z]*$") Then
						; Reset precheck array as required
						$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
						Return SetError(1, 2, "")
					EndIf
					; And store
					$sDrives =StringUpper($aSplit[1])
					; Check for default drive
					If $aSplit[0] > 1 Then
						$sDefDrive = StringUpper(StringLeft($aSplit[2], 1))
						; Check if default in drive list
						If $sDrives And (Not StringInStr($sDrives, $sDefDrive)) Then
							; Clear if not
							$sDefDrive = ""
						Else
							; Add colon
							If $sDefDrive Then
								$sDefDrive &= ":"
							EndIf
						EndIf
						; Is it a valid folder path?
						If FileExists($aSplit[2]) And StringInStr(FileGetAttrib($aSplit[2]), "D") Then
							$sDefFolder = $aSplit[2]
							; Check if trailing backslash
							If StringRight($sDefFolder, 1) = "\" Then
								; Set flag
								$bDefFolder_Open = True
								; Remove the trailing backslash
								$sDefFolder = StringTrimRight($sDefFolder, 1)
							EndIf
							; Replace backslashes for comparison when expanded
							$sDefFolder_Check = StringReplace($sDefFolder, "\", "|")
						EndIf
					EndIf
					; Clear root parameter to show drive choice required
					$sRoot = ""
				Else
					; Reset precheck array as required
					$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
					; Not valid path or drive list
					Return SetError(1, 1, "")
				EndIf
			EndIf
	EndSwitch

	If $sRoot = "" Then
		; Check if drive list set
		If $sDrives Then
			; Create array of these drives
			$aAll_Drives = StringSplit($sDrives, "")
			For $i = 1 To $aAll_Drives[0]
				$aAll_Drives[$i] &= ":"
			Next
		Else
			; Get array of drives
			$aAll_Drives = DriveGetDrive("ALL")
			If @error Then
				Local $aAll_Drives[1] = [0]
			EndIf
		EndIf
		; Get array of network drives - to show indexing label if required
		$aNetwork_Drives = DriveGetDrive("NETWORK")
		If @error Then
			Local $aNetwork_Drives[1] = [0]
		EndIf
	EndIf

	; Check Mask
	If $sMask = Default Then
		$sMask = "*"
	EndIf

	; Extract File mask (needed for hide extension code)
	Local $aMaskSplit = StringSplit($sMask, "|")
	Local $sFile_Mask = $aMaskSplit[1]

	; Check Display parameter
	$iDisplay = _CFF_Check_Display($iDisplay, $sFile_Mask, $bShow_Ext, $iHide_HS, $bBoth_Selectable, $bDuplicates_Allowed, $bFileScroll, $bSplash, $bDeepest)
	If @error Then
		; Reset precheck array as required
		$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
		Return SetError(3, 0, "")
	EndIf

	; Verify precheck list if set
	If IsArray($g_aCFF_PreCheck) Then
		_CFF_Verify_PreCheck($g_aCFF_PreCheck, $iDisplay, $bBoth_Selectable)
	EndIf

	; Check for return type and possible return ControlID
	Switch $cReturn
		Case 0
			; Single select
			$bSingle_Sel = True
			; Prevent from firing
			$cReturn = 9999
		Case Else
			; Multi select
			$bSingle_Sel = False
			; Check if valid ControlID
			If Not IsHWnd(GUICtrlGetHandle($cReturn)) Then
				; Prevent from firing
				$cReturn = 9999
			EndIf
	EndSwitch

	; If list passed check if valid ControlID
	If $cList And Not IsHWnd(GUICtrlGetHandle($cList)) Then
		; Reset precheck array as required
		$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
		Return SetError(5, 0, "")
	EndIf

	; Check if TreeView is using checkboxes
	If BitAnd(_WinAPI_GetWindowLong($g_hCFF_TreeView, 0xFFFFFFF0), 0x0100) Then ; $GWL_STYLE, $TVS_CHECKBOXES
		$bCheckBox = True
		; Check if folder checkboxes should be hidden - only files selectable and no prechecks
		If $iDisplay = 0 And Not($bBoth_Selectable) And Not(IsArray($g_aCFF_PreCheck)) Then
			$bNoFolderCheck = True
		EndIf
	EndIf

	; Create dummy controls
	$g_cCFF_Expand_Dummy = GUICtrlCreateDummy() ; fires when [+] clicked
	$g_cCFF_Select_Dummy = GUICtrlCreateDummy() ; fires when item checked
	$g_cCFF_Click_Dummy = GUICtrlCreateDummy()  ; Fires on Click
	$g_cCFF_DblClk_Dummy = GUICtrlCreateDummy() ; fires on DblClk
	Local $cRet_Dummy = GUICtrlCreateDummy()    ; fires on ENTER
	Local $cEsc_Dummy = GUICtrlCreateDummy()    ; fires on ESC
	Local $cRedraw_Dummy = GUICtrlCreateDummy() ; fires on Redraw key
	; Set accel keys
	Local $aAccelKeys[2][2] = [["{ENTER}", $cRet_Dummy],["{ESC}", $cEsc_Dummy]]
	GUISetAccelerators($aAccelKeys)

	; Fill tree
	If $sRoot Then
		; If root folder specified then fill TV
		$aFill_Ret = _CFF_Fill_Branch($vTV_ID, $cTreeView, $sRoot, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)
	Else
		_CFF_Fill_Drives($vTV_ID, $aAll_Drives, $sDefDrive, $sDefFolder, $bDeepest, $bNoFolderCheck)
	EndIf

	; Ensure first file visible if selecting files only and displaying both files and folders
	If (Not $bBoth_Selectable) And $aFill_Ret[2] Then
		; Last file visible
		_GUICtrlTreeView_EnsureVisible($g_hCFF_TreeView, $aFill_Ret[2])
		; Last folder visible
		_GUICtrlTreeView_EnsureVisible($g_hCFF_TreeView, $aFill_Ret[1])
	EndIf

	; Highlight passed file if required
	If $sDefFile Then
		$hItem = _GUICtrlTreeView_GetFirstItem($g_hCFF_TreeView)
		Do
			If _GUICtrlTreeView_GetText($g_hCFF_TreeView, $hItem) = $sDefFile Then
				_GUICtrlTreeView_EnsureVisible($g_hCFF_TreeView, $hItem)
				_GUICtrlTreeView_ClickItem($g_hCFF_TreeView, $hItem)
				ExitLoop
			EndIf
			$hItem = _GUICtrlTreeView_GetNextChild($g_hCFF_TreeView, $hItem)
		Until $hItem = 0
	EndIf

	; If defined folder
	If $sDefFolder Then
		; Fire auto expansion in loop
		GUICtrlSendToDummy($cRedraw_Dummy)
		; Set counter
		$iRedraw_Count = 0
	EndIf

	If $bCheckBox Then
		; Parse TV for check data
		$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
		$aOldCheckData = $aTVCheckData
	EndIf

	; Change to MessageLoop mode
	Local $nOldOpt = Opt('GUIOnEventMode', 0)

	; Clear flag for Ret key pressed
	Local $iRet_Pressed = 0

	While 1

		Switch GUIGetMsg()

			Case $cRedraw_Dummy
				; Empty TreeView
				_GUICtrlTreeView_DeleteAll($g_hCFF_TreeView)
				; Refill drives
				_CFF_Fill_Drives($g_hCFF_TreeView, $aAll_Drives, $sDefDrive, $sDefFolder, $bDeepest, $bNoFolderCheck)
				; Increase counter
				$iRedraw_Count += 1
				If $iRedraw_Count < 5 Then
					; Re-expand to default folder
					_CFF_AutoExpand($g_hCFF_TreeView, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $sDefFolder, $bDefFolder_Open, $iRedraw_Count, $bDeepest, $bNoFolderCheck)
					; Check expansion correct
					If _GUICtrlTreeView_GetTree($g_hCFF_TreeView) <> $sDefFolder_Check Then
						; Force another redraw
						GUICtrlSendToDummy($cRedraw_Dummy)
					EndIf
				EndIf
				; Reparse TV
				If $bCheckBox Then
					$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
				EndIf

			Case $g_cCFF_Click_Dummy

				If $bCheckBox Then
					; Loop through current TreeView items to see if a checkbox has changed state
					; Get handle of selected item
						If $bNative_TV Then
							$hItem = GUICtrlGetHandle(GUICtrlRead($cTreeView))
						Else
							$hItem = _GUICtrlTreeView_GetSelection($g_hCFF_TreeView)
						EndIf
						For $i = 0 To UBound($aOldCheckData) - 1
						; Compare handle to existing item
						If $hItem <> $aOldCheckData[$i][0] Then ExitLoop ; TreeView was obviously expanded
						; Compare item checkbox states
						If _GUICtrlTreeView_GetChecked($g_hCFF_TreeView, $hItem) <> $aOldCheckData[$i][1] Then
							; Determine if item path is a valid selection
							$sSelectedPath = _CFF_Check_Selection($g_hCFF_TreeView, $hItem, $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
							; If not valid selection
							If Not $sSelectedPath Then
								; If prechecks exist
								If IsArray($g_aCFF_PreCheck) Then
									; Force to that state
									_GUICtrlTreeView_SetChecked($cTreeView, $hItem, $aOldCheckData[$i][1])
								Else
									; If folder checkboxes are shown
									If Not($bNoFolderCheck) Then
										; Force to unchecked
										_GUICtrlTreeView_SetChecked($cTreeView, $hItem, False)
									EndIf
								EndIf
							EndIf
							; No point in looking further so check for parent/child adjustments
							GUICtrlSendToDummy($g_cCFF_Select_Dummy, $hItem)
							ExitLoop
						EndIf
						; Get next item handle
						$hItem = _GUICtrlTreeView_GetNext($g_hCFF_TreeView, $hItem)
						; Exit if TreeView had been contracted
						If $hItem = 0 Then ExitLoop
					Next
				EndIf

			Case $g_cCFF_Select_Dummy

				If $bCheckBox Then
					; Check if Control or Alt pressed
					_WinAPI_GetAsyncKeyState(0x11) ; Needed to avoid double setting
					_WinAPI_GetAsyncKeyState(0x12)
					If _WinAPI_GetAsyncKeyState(0x11) Or _WinAPI_GetAsyncKeyState(0x12) Then
						; Get handle of selected item
						$hItem = GUICtrlRead($g_cCFF_Select_Dummy)
						; Determine checked state
						Local $bState = _GUICtrlTreeView_GetChecked($cTreeView, $hItem)
						; Find item in array
						Local $iItemIndex = _ArraySearch($aOldCheckData, $hItem)
						; If checked state has altered
						If $aOldCheckData[$iItemIndex][1] <> $bState Then
							; Store new state
							$aOldCheckData[$iItemIndex][1] = $bState
							; Adjust parents if only deepest item is to be returned and Alt pressed
							If $bDeepest And _WinAPI_GetAsyncKeyState(0x12) Then
								_CFF_Adjust_Parents($g_hCFF_TreeView, $hItem, $aOldCheckData, $bState)
							EndIf
							; Adjust visible children if Ctrl pressed
							If _WinAPI_GetAsyncKeyState(0x11) Then
								_CFF_Adjust_Children($cTreeView, $hItem, $aOldCheckData, $bState)
								; Expand item if required
								If Not _GUICtrlTreeView_GetExpanded($cTreeView, $hItem) Then
									GUICtrlSendToDummy($g_cCFF_Expand_Dummy, $hItem)
								EndIf
							EndIf
						EndIf
					EndIf
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
					$aOldCheckData = $aTVCheckData
				EndIf

			Case $cRet_Dummy
				; Set flag
				$iRet_Pressed = 1
				If $bSingle_Sel Then
					; Get selected item
					If $bNative_TV Then
						$hItem = GUICtrlGetHandle(GUICtrlRead($cTreeView))
					Else
						$hItem = _GUICtrlTreeView_GetSelection($g_hCFF_TreeView)
					EndIf
					If $hItem Then
						; Determine item path
						$sAddFile_List = $sRoot & StringReplace(_GUICtrlTreeView_GetTree($g_hCFF_TreeView, $hItem), "|", "\")
						Switch $iDisplay
							Case 2
								; Folders only
							Case 3
								; Check if any children
								If _GUICtrlTreeView_GetChildCount($g_hCFF_TreeView, $hItem) < 1 Then
									; No children so selection permitted
									ContinueCase
								EndIf
							Case Else
								; Files only
								StringReplace(FileGetAttrib($sAddFile_List), "D", "")
								; Is it a folder?
								If @extended Then
									$sAddFile_List = ""
								EndIf
								; Hide extension?
								If $bShow_Ext = False Then
									$sAddFile_List &= StringTrimLeft($sFile_Mask, 1)
								EndIf
						EndSwitch
					EndIf
				EndIf
				ContinueCase

			Case $cReturn
				If $bCheckBox Then
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
					$aOldCheckData = $aTVCheckData
					; Construct return list
					Local $sOldSep = Opt("GUIDataSeparatorChar", "\"), $sFullPath
					$sAddFile_List = ""
					For $i = 0 To UBound($aTVCheckData) - 1
						If $aTVCheckData[$i][1] And $aTVCheckData[$i][2] Then
							; Check if valid selection
							$sFullPath = _CFF_Check_Selection($g_hCFF_TreeView, $aTVCheckData[$i][0], $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
							If $sFullPath Then $sAddFile_List &= $sFullPath & "|"
						EndIf
					Next
					Opt("GUIDataSeparatorChar", $sOldSep)
					; Check if only deepest items are to be returned
					If $bDeepest Then
						; Convert list to array
						Local $aDeepest = StringSplit($sAddFile_List, "|")
						; Loop through array checking if item forms part of preceding item
						For $i = $aDeepest[0] - 2 To 1 Step -1
							$sFullPath = $aDeepest[$i]
							If StringInStr(FileGetAttrib($sFullPath), "D") Then
								If StringInStr($aDeepest[$i + 1], $aDeepest[$i]) Then
									; Remove item from list
									$sAddFile_List = StringReplace($sAddFile_List, $aDeepest[$i] & "|", "")
								EndIf
							EndIf
						Next
						; If there were prechecked items
						If IsArray($g_aCFF_PreCheck) Then
							; Loop through to remove any parent items that were not fully expanded
							For $i = 1 To $aDeepest[0]
								For $j = 0 To UBound($g_aCFF_PreCheck) - 1
									If StringInStr($g_aCFF_PreCheck[$j], $aDeepest[$i]) Then
										$sAddFile_List = StringReplace($sAddFile_List, $aDeepest[$i] & "|", "")
									EndIf
								Next
							Next
						EndIf
					EndIf
					; Now add unused prechecked items
					For $i = 0 To UBound($g_aCFF_PreCheck) - 1
						If $g_aCFF_PreCheck[$i] Then
							$sAddFile_List &= $g_aCFF_PreCheck[$i] & "|"
						EndIf
					Next
				EndIf
				; Clear treeview and list
				_GUICtrlTreeView_DeleteAll($g_hCFF_TreeView)
				GUICtrlSetData($cList, "|")
				; Cancel Accel keys
				GUISetAccelerators(0)
				; Reset precheck array as required
				$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
				; Return list of files selected
				Return SetError(0, $iRet_Pressed, StringTrimRight($sAddFile_List, 1))

			Case $g_cCFF_DblClk_Dummy
				If $bCheckBox Then
					; Reparse TV
					$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
				Else
					; Clear flag
					Local $bPermit = False
					; Check doubleclick selection is permitted (not default for folders only)
					If $iDisplay = 2 Then
						; Not permitted
					ElseIf $iDisplay = 3 Then
						; Check if any children
						If $bNative_TV Then
							$hItem = GUICtrlGetHandle(GUICtrlRead($cTreeView))
						Else
							$hItem = _GUICtrlTreeView_GetSelection($g_hCFF_TreeView)
						EndIf
						If _GUICtrlTreeView_GetChildCount($g_hCFF_TreeView, $hItem) < 1 Then
							; No children so selection permitted
							$bPermit = True
						EndIf
					Else
						; Permitted
						$bPermit = True
					EndIf
					; If selection permitted
					If $bPermit Then
						; Get item data
						If $bNative_TV Then
							$hItem = GUICtrlGetHandle(GUICtrlRead($cTreeView))
						Else
							$hItem = _GUICtrlTreeView_GetSelection($g_hCFF_TreeView)
						EndIf
						; Check path is a valid selection
						$sSelectedPath = _CFF_Check_Selection($g_hCFF_TreeView, $hItem, $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)
						If $sSelectedPath Then
							; Valid item selected
							If $bSingle_Sel Then
								; Clear treeview
								_GUICtrlTreeView_DeleteAll($g_hCFF_TreeView)
								; Restore previous mode
								Opt('GUIOnEventMode', $nOldOpt)
								; Cancel Accel keys
								GUISetAccelerators(0)
								; Reset precheck array as required
								$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
								; Return valid path
								Return $sSelectedPath
							Else
								; Check if Ctrl pressed
								_WinAPI_GetAsyncKeyState(0x11) ; Needed to avoid double setting
								If _WinAPI_GetAsyncKeyState(0x11) Then
									$sAddFile_List = _CFF_List_Del($sAddFile_List, $sSelectedPath, $cList, $g_hCFF_TreeView)
								Else
									$sAddFile_List = _CFF_List_Add($sAddFile_List, $sSelectedPath, $cList, $bDuplicates_Allowed, $g_hCFF_TreeView)
								EndIf
							EndIf
						EndIf
					EndIf
				EndIf

			Case $g_cCFF_Expand_Dummy
				; Get expanded item hamdle
				$hItem = GUICtrlRead($g_cCFF_Expand_Dummy)
				If $hItem Then
					; Select item
					_GUICtrlTreeView_ClickItem($g_hCFF_TreeView, $hItem)
					; Get expanded item data
					If $bNative_TV Then
						$cItem = GUICtrlRead($cTreeView)
						$hItem = GUICtrlGetHandle($cItem)
					Else
						$hItem = _GUICtrlTreeView_GetSelection($g_hCFF_TreeView)
						$cItem = $hItem
					EndIf
					$sSelectedPath = $sRoot & StringReplace(_GUICtrlTreeView_GetTree($g_hCFF_TreeView, $hItem), "|", "\")
					; Check if dummy child exists or has already been filled
					Local $hFirstChild = _GUICtrlTreeView_GetFirstChild($g_hCFF_TreeView, $hItem)
					Local $sFirstChild = _GUICtrlTreeView_GetText($g_hCFF_TreeView, $hFirstChild)
					; If dummy child exists
					If $sFirstChild = "" Then
						; Fill with content
						$aFill_Ret = _CFF_Fill_Branch($vTV_ID, $cItem, $sSelectedPath, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)
						; Delete the dummy
						_GUICtrlTreeView_Delete($g_hCFF_TreeView, $hFirstChild)
					EndIf
					; If file scroll selected AND files were found AND the branch is being expanded
					If $bFileScroll And $aFill_Ret[2] And _GUICtrlTreeView_GetExpanded($g_hCFF_TreeView, $hItem) Then
						; Scroll to first file in this branch
						_CFF_File_Visible($g_hCFF_TreeView, $aFill_Ret)
					EndIf
					; Clear the flag to reactivate the handler
					GUICtrlSendToDummy($g_cCFF_Expand_Dummy, 0)
					If $bCheckBox Then
						; Reparse TV
						$aTVCheckData = _CFF_ParseTV($g_hCFF_TreeView)
						; Save current state
						$aOldCheckData = $aTVCheckData
					EndIf
				EndIf

			Case $cEsc_Dummy
				; Clear treeview and list
				_GUICtrlTreeView_DeleteAll($g_hCFF_TreeView)
				GUICtrlSetData($cList, "|")
				; Restore previous mode
				Opt('GUIOnEventMode', $nOldOpt)
				; Cancel Accel keys
				GUISetAccelerators(0)
				; Reset precheck array as required
				$g_aCFF_PreCheck = $g_aCFF_PreCheckRetain
				; Return
				Return SetError(7, 0, "")

		EndSwitch

	WEnd

EndFunc   ;==>_CFF_Embed

; #FUNCTION# =========================================================================================================
; Name...........: _CFF_SetPreCheck
; Description ...: Sets a list of files/folders which will be checked when expanded
; Syntax.........: _CFF_SetPreCheck ($aPreCheck_List [, $bNoPartial = True]))
; Parameters ....: $aPreCheck_List - Array holding paths of files/folders to check on expansion (or non-array to clear)
;                  $bNoPartial     - True (default) = partial paths removed from list
;                                    False = List untouched
;                  $bVolatile      - True (default) = List cleared after call to _CFF_Choose/Embed
;                                    False =  List remains active for future calls
; Requirement(s).: v3.3 +
; Return values .: 1 = Precheck list set
;                  0 = Precheck list cleared
;                  @extended = state of $g-bCFF_PreCheckVolatile flag
; Author ........: Melba23
; Modified ......:
; Remarks .......: - Items will be checked when branch containing the item is expanded.  If the item has not been expanded
;                    it will be automatically included in the return list.
;                  - The UDF will remove any invalid paths and any that would not be displayed in the tree because of the
;                    display setting
;                  - Folder paths can be with/without a trailing \ which is automatically added
;                  - Passing a non-array parameter clears pre-check list
;                  - Partial paths are part of other paths within the array.  If the "+ 512" (only deepest element returned)
;                    option is used these paths need to be removed from the array
; Example........: Yes
;=====================================================================================================================
Func _CFF_SetPreCheck($aPreCheck_List, $bNoPartial = True, $bVolatile = True)

	If Not IsArray($aPreCheck_List) Then

		; Clear list
		$aPreCheck_List = ""

	Else

		Local $sTest, $sTestSRE, $bFolder, $iMax = UBound($aPreCheck_List) - 1

		If $bNoPartial = Default Then $bNoPartial = True
		If $bVolatile = Default Then $bVolatile = True

		; Loop through elements
		For $i = 0 To $iMax
			; Extract path
			$sTest = $aPreCheck_List[$i]

			; Check for drives
			If StringRegExp($sTest, "^[A-Za-z]:\\?$") Then
				; Check if valid (add trailing "\" if required) and delete if not
				If DriveStatus($sTest & ((StringRight($sTest, 1) = "\") ? ("") : ("\"))) <> "READY" Then
					$aPreCheck_List[$i] = ""
					; No point in checking further
					ContinueLoop
				EndIf
			EndIf

			; Check for partial paths if NoPartial set
			If $bNoPartial Then
				; Check if folder
				If StringInStr(FileGetAttrib($sTest), "D") Then
					; Force a trailing "\"
					If StringRight($sTest, 1) <> "\" Then
						$sTest &= "\"
						$aPreCheck_List[$i] &= "\"
					EndIf
					; Set flag
					$bFolder = True
					; Escape any SRE special characters
					$sTestSRE = StringRegExpReplace($sTest, "[][$^.{}()+\\-]", "\\$0")
				Else
					; Clear flag
					$bFolder = False
				EndIf
				; Compare to all other elements
				For $j = 0 To $iMax
					; If not test element itself AND is folder - check for partial path
					If $j <> $i And $bFolder And StringRegExp($aPreCheck_List[$j], "^" & $sTestSRE) Then
						; And delete if so
						$aPreCheck_List[$i] = ""
						; No point in checking further
						ExitLoop
					EndIf
				Next
			EndIf
		Next

	EndIf

	; Store final array
	$g_aCFF_PreCheck = $aPreCheck_List
	; Check for precheck volatility
	If $bVolatile Then
		$g_aCFF_PreCheckRetain = ""
	Else
		$g_aCFF_PreCheckRetain = $aPreCheck_List
	EndIf

	; Return action code
	Return SetExtended(IsArray($g_aCFF_PreCheckRetain), IsArray($g_aCFF_PreCheck))

EndFunc

; #FUNCTION# =========================================================================================================
; Name...........: _CFF_RegMsg
; Description ...: Registers WM_NOTIFY for TreeView doubleclick and item expansion and WM_COMMAND for list focus
; Syntax.........: _CFF_RegMsg([$bNOTIFY = True[, $bCOMMAND = True]])
; Parameters.....: $bNOTIFY  - True (default) = Register WM_NOTIFY handler function
;                  $bCOMMAND - True (default) = Register WM_COMMAND handler function
; Requirement(s).: v3.3 +
; Return values .: 0 - No handlers registered
;                  1 - WM_NOTIFY handler registered
;                  2 - WM_COMMAND handler registered
;                  3 - Both handlers registerd
; Author ........: Melba23
; Modified ......:
; Remarks .......: If a script has existing WM_NOTIFY or WM_COMMAND handlers then do not register the relevant handler
;                  using this function but call it directly from within the existing message handler
; Example........: Yes
;=====================================================================================================================
Func _CFF_RegMsg($bNOTIFY = True, $bCommand = True)

	; Register required messages
	Local $iRet = 0
	If $bNOTIFY Then
		$iRet = GUIRegisterMsg(0x004E, "_CFF_WM_NOTIFY_Handler") ; $WM_NOTIFY
	EndIf
	If $bCommand Then
		$iRet = 2 * GUIRegisterMsg(0x0111, "_CFF_WM_COMMAND_Handler") ; $WM_COMMAND
	EndIf
	Return $iRet

EndFunc   ;==>_CFF_RegMsg

; #FUNCTION# =========================================================================================================
; Name...........: _CFF_WM_NOTIFY_Handler
; Description ...: Windows message handler for WM_NOTIFY - reacts to doubleclick and item expansion on TreeView
; Syntax.........: _CFF_WM_NOTIFY_Handler($hWnd, $iMsg, $wParam, $lParam)
; Requirement(s).: v3.3 +
; Return values..: None
; Author ........: Melba23 - thanks to guinness for help with the #*@#*%# Struct !!!!
; Modified ......:
; Remarks .......: If a WM_NOTIFY handler already registered, then call this function from within that handler
; Example........: Yes
;=====================================================================================================================
Func _CFF_WM_NOTIFY_Handler($hWnd, $iMsg, $wParam, $lParam)

	#forceref $hWnd, $iMsg, $wParam

	; Create NMTREEVIEW structure
	Local $tStruct = DllStructCreate("struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct;" & _
			"uint Action;struct;uint OldMask;handle OldhItem;uint OldState;uint OldStateMask;" & _
			"ptr OldText;int OldTextMax;int OldImage;int OldSelectedImage;int OldChildren;lparam OldParam;endstruct;" & _
			"struct;uint NewMask;handle NewhItem;uint NewState;uint NewStateMask;" & _
			"ptr NewText;int NewTextMax;int NewImage;int NewSelectedImage;int NewChildren;lparam NewParam;endstruct;" & _
			"struct;long PointX;long PointY;endstruct", $lParam)
	Local $hWndFrom = DllStructGetData($tStruct, "hWndFrom")
	Local $hItem = DllStructGetData($tStruct, "NewhItem")
	Local $iCode = DllStructGetData($tStruct, "Code")

	If $hWndFrom = $g_hCFF_TreeView Then
		; Set flag for TreeView actioned
		$g_bCFF_ActiveTV = True
		; Clear list selection
		DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $g_hCFF_List, "uint", 0x0186, "wparam", -1, "lparam", 0) ; $LB_SETCURSEL
		; Check action
		Switch $iCode
			Case -2 ; $NM_CLICK
				; Fire the dummy control
				GUICtrlSendToDummy($g_cCFF_Click_Dummy)
			Case -3 ; $NM_DBLCLK 0xFFFFFFFD
				; Fire the dummy control
				GUICtrlSendToDummy($g_cCFF_DblClk_Dummy)
				; Set flag
			Case $TVN_ITEMEXPANDEDW, $TVN_ITEMEXPANDEDA
				; Check autoexpansion flag is not set
				If Not $g_bCFF_AutoExpand Then
					; Fire the dummy control if expanding
					If DllStructGetData($tStruct, "Action") = 2 Then
						GUICtrlSendToDummy($g_cCFF_Expand_Dummy, $hItem)
					EndIf
				EndIf
			Case $TVN_SELCHANGEDA, $TVN_SELCHANGEDW
				; Fire the dummy control
				GUICtrlSendToDummy($g_cCFF_Select_Dummy, $hItem)
		EndSwitch
	EndIf

EndFunc   ;==>_CFF_WM_NOTIFY_Handler

; #FUNCTION# =========================================================================================================
; Name...........: _CFF_WM_COMMAND_Handler
; Description ...: Windows message handler for WM_COMMAND - reacts to list getting focus
; Syntax.........: _CFF_WM_COMMAND_Handler($hWnd, $iMsg, $wParam, $lParam)
; Requirement(s).: v3.3 +
; Return values..: None
; Author ........: Melba23
; Modified ......:
; Remarks .......: If a WM_COMMAND handler already registered, then call this function from within that handler
; Example........: Yes
;=====================================================================================================================
Func _CFF_WM_COMMAND_Handler($hWnd, $iMsg, $wParam, $lParam)

	#forceref $hWnd, $iMsg

	; If list actioned
	If $lParam = $g_hCFF_List Then
		; If item selected
		If BitShift($wParam, 16) = 0x0001 Then ; $LBN_SELCHANGE
			; Clear flag
			$g_bCFF_ActiveTV = False
		EndIf
	EndIf

EndFunc   ;==>_CFF_WM_COMMAND_Handler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Combo_Fill
; Description ...: Creates and fills a combo for drive selection.
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_Fill_Combo( $iW, $sDrives, ByRef $sDefDrive)

	Local $iInset = Int(($iW - 125) / 2), $aDrives

	; Create drive array
	If $sDrives Then
		; Use specified drive list
		$aDrives = StringSplit($sDrives, "")
		For $i = 1 To $aDrives[0]
			$aDrives[$i] &= ":"
		Next
	Else
		; If no drives specified then list all
		$aDrives = DriveGetDrive("ALL")
	EndIf

	; Create drive list for combo
	$sDrives = ""
	For $i = 1 To $aDrives[0]
		; Only display ready drives
		If DriveStatus($aDrives[$i] & '\') == "READY" Then
			$sDrives &= StringUpper($aDrives[$i]) & "|"
		Else
			; If default drive not ready
			If $sDefDrive = $aDrives[$i] Then
				; Clear it
				$sDefDrive = ""
			EndIf
		EndIf
	Next

	; Create combo
	GUICtrlCreateLabel("Select Drive:", $iInset, 15, 65, 20)
	Local $cCombo = GUICtrlCreateCombo("", 65 + $iInset, 10, 50, 20)
	GUICtrlSetData($cCombo, $sDrives, $sDefDrive)
	Return $cCombo

EndFunc   ;==>_CFF_Combo_Fill

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Fill_Drives
; Description ...: Fills a TreeView with ready drives
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_Fill_Drives($cTV, $aAll_Drives, $sDefDrive, $sDefFolder, $bDeepest, $bNoFolderCheck)

	Local $hTV, $cItem

	; Check if native or UDF TV
	Local $bNative_TV = True
	If IsHWnd($cTV) Then
		$bNative_TV = False
		$hTV = $cTV
	Else
		$hTV = GUICtrlGetHandle($cTV)
	EndIf

	Local $sDrive, $cDefDriveItem = 0, $hItem
	_GUICtrlTreeView_BeginUpdate($hTV)
	For $i = 1 To $aAll_Drives[0]
		; Extract drive
		$sDrive = $aAll_Drives[$i]
		; Only display ready drives
		If DriveStatus($sDrive & '\') == "READY" Then
			If $bNative_TV Then
				$cItem = GUICtrlCreateTreeViewItem(StringUpper($sDrive), $cTV)
				GUICtrlCreateTreeViewItem("", $cItem)
				If $sDrive = $sDefDrive Then
					$cDefDriveItem = $cItem * - 1
				EndIf
				$hItem = GUICtrlGetHandle($cItem)
			Else
				$hItem = _GUICtrlTreeView_Add($hTV, $cTV, StringUpper($sDrive))
				_GUICtrlTreeView_AddChild($hTV, $hItem, "")
				If $sDrive = $sDefDrive Then
					$cDefDriveItem = $cItem
				EndIf
			EndIf
		EndIf

		; Check if only files selectable and no precheck
		If $bNoFolderCheck Then
			; Hide folder checkboxes
			_GUICtrlTreeView_SetStateImageIndex($hTV, $hItem, 0)
		EndIf

		; Check for PreCheck drives
		If IsArray($g_aCFF_PreCheck) Then
			; Loop through array
			For $j = 0 To UBound($g_aCFF_PreCheck) - 1
				; Check if item is part of a precheck path
				If StringRegExp($g_aCFF_PreCheck[$j], "(?i)^" & $sDrive & "\\") Then
					; Set checked if deepest flag set
					If $bDeepest Then
						_GUICtrlTreeView_SetChecked($hTV, $hItem)
					EndIf
					; Check for complete match
					If StringRegExp($g_aCFF_PreCheck[$j], "(?i)^" & $sDrive & "\\$") Then
						; If so then check and remove from list
						_GUICtrlTreeView_SetChecked($hTV, $hItem)
						$g_aCFF_PreCheck[$j] = ""
					EndIf
					; No point in looking further
					ExitLoop
				EndIf
			Next
		EndIf

	Next

	; Expand default drive if set - but not if definded folder also set
	If $cDefDriveItem And Not $sDefFolder Then
		If $cDefDriveItem < 0 Then
			; Convert ControlID to handle
			$cDefDriveItem = GUICtrlGetHandle($cDefDriveItem * -1)
		EndIf
		GUICtrlSendToDummy($g_cCFF_Expand_Dummy, $cDefDriveItem)
		_GUICtrlTreeView_Expand($hTV, $cDefDriveItem)
	EndIf
	_GUICtrlTreeView_EndUpdate($hTV)

EndFunc   ;==>_CFF_Fill_Drives

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Fill_Branch
; Description ...: Fills a TreeView branch with folders and files on expansion
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_Fill_Branch($cTV, $cParent, $sPath, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)

	Local $hTV, $sFolder_Mask, $sItem, $cItem, $hSearch, $aFill_Ret[3] = [False, 0, 0], $sTree, $aContent

	; Check if network drive
	Local $sDrive = StringLeft($sPath, 2)
	If IsArray($aNetwork_Drives) Then
		For $i = 1 To $aNetwork_Drives[0]
			If $aNetwork_Drives[$i] = $sDrive Then
				Local $aTV_Pos = WinGetPos(GUICtrlGetHandle($cTV))
				SplashOff() ; Close any existing splash screen
				SplashTextOn("", "Indexing from" & @CRLF & "Network drive..." & @CRLF & @CRLF & "Please be patient", $aTV_Pos[2], $aTV_Pos[3], $aTV_Pos[0], $aTV_Pos[1], 33)
			EndIf
		Next
	EndIf

	; Check if native or UDF TV
	Local $bNative_TV = True
	If IsHWnd($cTV) Then
		$bNative_TV = False
		$hTV = $cTV
	Else
		$hTV = GUICtrlGetHandle($cTV)
	EndIf

	; Store parent handle for eventual scroll if only files found
	If IsHWnd($cParent) Then
		$aFill_Ret[1] = $cParent
	Else
		$aFill_Ret[1] = GUICtrlGetHandle($cParent)
	EndIf

	; Force path with trailing \
	If StringRight($sPath, 1) <> "\" Then
		$sPath &= "\"
	EndIf

	; Change seperator character for possible tree returns
	Local $sOldSep = Opt("GUIDataSeparatorChar", "\")

	; Expand TV
	_GUICtrlTreeView_BeginUpdate($hTV)
	; Search for folders if required
	Switch $iDisplay
		Case 0, 2, 3
			; Set mask for folder search
			$sFolder_Mask = $sMask ; This is valid for folder only search
			If $iDisplay = 0 Then ; For file and folder search look for a folder_Exclude parameter
				Local $aMaskSplit = StringSplit($sMask, "|")
				If $aMaskSplit[0] = 3 Then
					; And use it if found
					$sFolder_Mask = "*|" & $aMaskSplit[3]
				Else
					$sFolder_Mask = "*"
				EndIf
			EndIf
			; Now list folders
			$aContent = _FileListToArrayRec($sPath, $sFolder_Mask, 2 + $iHide_HS, 0, 1, 1)
			If IsArray($aContent) Then
				For $i = 1 To $aContent[0]
					; Remove trailing \ if needed
					$sItem = $aContent[$i]
					If StringRight($sItem, 1) = "\" Then
						$sItem = StringTrimRight($sItem, 1)
					EndIf
					; Create item
					If $bNative_TV Then
						$cItem = GUICtrlCreateTreeViewItem($sItem, $cParent)
						; Store handle for possible scroll
						$aFill_Ret[1] = GUICtrlGetHandle($cItem)
					Else
						; Use correct function depending on TV type
						If $cParent = $hTV Then
							$cItem = _GUICtrlTreeView_Add($hTV, $cParent, $sItem)
						Else
							$cItem = _GUICtrlTreeView_AddChild($hTV, $cParent, $sItem)
						EndIf
						; Store handle for possible scroll
						$aFill_Ret[1] = $cItem
					EndIf

					; Check if only files selectable and no precheck
					If $bNoFolderCheck Then
						; Hide folder checkboxes
						_GUICtrlTreeView_SetStateImageIndex($hTV, $aFill_Ret[1], 0)
					EndIf

					; Check for PreCheck folders
					If IsArray($g_aCFF_PreCheck) Then
						; Get full tree of the item
						$sTree = $sPath & StringRegExpReplace(_GUICtrlTreeView_GetTree($hTV, $aFill_Ret[1]), ".*\\(.+)", "$1")
						; Escape special SRE characters
						$sTree = StringRegExpReplace($sTree, "[][$^.{}()+\\-]", "\\$0")
						; Loop through array
						For $j = 0 To UBound($g_aCFF_PreCheck) - 1
							; Check if item is part of a precheck path
							If StringRegExp($g_aCFF_PreCheck[$j], "(?i)^" & $sTree & "\\") Then
								; Set checked if deepest flag set
								If $bDeepest Then
									_GUICtrlTreeView_SetChecked($hTV, $aFill_Ret[1])
								EndIf
								; Check for complete match
								If StringRegExp($g_aCFF_PreCheck[$j], "(?i)^" & $sTree & "\\$") Then
									; If so then check and remove from list
									_GUICtrlTreeView_SetChecked($hTV, $aFill_Ret[1])
									$g_aCFF_PreCheck[$j] = ""
								EndIf
								; No point in looking further
								ExitLoop
							EndIf
						Next
					EndIf

					; Force a [+] if suitable content within
					FileChangeDir($sPath & $aContent[$i])
					$hSearch = FileFindFirstFile("*")
					; If there is something within
					If $hSearch <> -1 Then
						; Set flag
						$aFill_Ret[0] = True
						; Different display modes require different content types
						Switch $iDisplay
							Case 0 ; Either folder or file
								; Create dummy child to force [+] display
								If $bNative_TV Then
									GUICtrlCreateTreeViewItem("", $cItem)
								Else
									_GUICtrlTreeView_AddChild($hTV, $cItem, "")
								EndIf
								; Case 1
								; Files only so no requirement for [+]
							Case 2, 3 ; Folder only
								While 1
									; Search for a folder
									FileFindNextFile($hSearch)
									; End of content
									If @error Then
										ExitLoop
									EndIf
									; If folder found
									If @extended Then
										; Create dummy child to force [+] display
										If $bNative_TV Then
											GUICtrlCreateTreeViewItem("", $cItem)
										Else
											_GUICtrlTreeView_AddChild($hTV, $cItem, "")
										EndIf
										; No need to look further
										ExitLoop
									EndIf
								WEnd
								; Close search
								FileClose($hSearch)
						EndSwitch
					EndIf
					; Reset working folder
					FileChangeDir(@ScriptDir)
				Next
			EndIf
	EndSwitch
	; Search for files if required
	Switch $iDisplay
		Case 0, 1
			; List files
			$aContent = _FileListToArrayRec($sPath, $sMask, 1 + $iHide_HS, 0, 1)
			If IsArray($aContent) Then
				; Set flag
				$aFill_Ret[0] = True
				For $i = 1 To $aContent[0]
					$sItem = $aContent[$i]
					; Remove extension if required
					If Not $bShow_Ext Then
						$sItem = StringRegExpReplace($sItem, "(.*)\..*", "$1")
					EndIf
					; Create item
					If $bNative_TV Then
						$cItem = GUICtrlCreateTreeViewItem($sItem, $cParent)
						; Store handle for possible scroll
						$aFill_Ret[2] = GUICtrlGetHandle($cItem)
					Else
						; Use correct function depending on TV type
						If $cParent = $hTV Then
							$cItem = _GUICtrlTreeView_Add($hTV, $cParent, $sItem)
						Else
							$cItem = _GUICtrlTreeView_AddChild($hTV, $cParent, $sItem)
						EndIf
						; Store handle for possible scroll
						$aFill_Ret[2] = $cItem
					EndIf

					; Check for PreCheck files
					If IsArray($g_aCFF_PreCheck) Then
						; Get full tree of the item
						$sTree = $sPath & StringRegExpReplace(_GUICtrlTreeView_GetTree($hTV, $aFill_Ret[2]), ".*\\(.+)", "$1")
						; Loop through the list
						For $j = 0 To UBound($g_aCFF_PreCheck) - 1
							; If there is a match
							If $g_aCFF_PreCheck[$j] = $sTree Then
								; Set checkbox
								_GUICtrlTreeView_SetChecked($hTV, $aFill_Ret[2])
								; Delete from list
								$g_aCFF_PreCheck[$j] = ""
								; No point in continuing
								ExitLoop
							EndIf
						Next
					EndIf

				Next
			EndIf
	EndSwitch
	_GUICtrlTreeView_EndUpdate($hTV)

	; Reset separator
	Opt("GUIDataSeparatorChar", $sOldSep)

	; Hide Splash if used
	SplashOff()

	Return $aFill_Ret

EndFunc   ;==>_CFF_Fill_Branch

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_AutoExpand
; Description ...: Expand tree to defined folder on start
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_AutoExpand($hTreeView, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $sDefFolder, $bDefFolder_Open, $iRedraw_Count, $bDeepest, $bNoFolderCheck, $bCombo = False, $bNative_TV = False)

	Local $aDefFolder_Split, $iDefFolder_Start = 1, $hDefFolder_Item, $vParent_Item, $sDefExpand_Path = ""

	; Set flag to prevent expansion dummy firing
	$g_bCFF_AutoExpand = True

	; Split path into elements
	$aDefFolder_Split = StringSplit($sDefFolder, "\")
	; Set up counter - miss drive element if not in treeview
	If $bCombo Then
		$iDefFolder_Start = 2
		$sDefExpand_Path = StringReplace($aDefFolder_Split[1], "_", " ") & "\"
	EndIf
	; Get initial handle to start process
	$hDefFolder_Item = _GUICtrlTreeView_GetFirstItem($hTreeView)
	; Look for required folder
	For $i = $iDefFolder_Start To $aDefFolder_Split[0]
		; Set required path
		$sDefExpand_Path &= StringReplace($aDefFolder_Split[$i], "_", " ") & "\"
		; Look for required folder
		Do
			If _GUICtrlTreeView_GetText($hTreeView, $hDefFolder_Item) = $aDefFolder_Split[$i] Then
				ExitLoop
			EndIf
			$hDefFolder_Item = _GUICtrlTreeView_GetNextSibling($hTreeView, $hDefFolder_Item)
		Until $hDefFolder_Item = 0
		;  Ensure visible and selected
		_GUICtrlTreeView_EnsureVisible($hTreeView, $hDefFolder_Item)
		_GUICtrlTreeView_ClickItem($hTreeView, $hDefFolder_Item)
		; Check if final folder is to be expanded
		If (Not $bDefFolder_Open) And $i = $aDefFolder_Split[0] Then
			ExitLoop
		EndIf
		; Expand folder
		_GUICtrlTreeView_Expand($hTreeView, $hDefFolder_Item)
		; Get ControlID/handle for parent
		If $bNative_TV Then
			; Short sleep for native ListViews internal processing - increasing with redraws
			Sleep(50 * $iRedraw_Count)
			$vParent_Item = GUICtrlRead($hTreeView)
		Else
			$vParent_Item = $hDefFolder_Item
		EndIf
		; Fill branch
		_CFF_Fill_Branch($hTreeView, $vParent_Item, $sDefExpand_Path, $iDisplay, $sMask, $bShow_Ext, $iHide_HS, $aNetwork_Drives, $bDeepest, $bNoFolderCheck)
		; Check for content - final folder may be empty
		If _GUICtrlTreeView_GetChildCount($hTreeView, $hDefFolder_Item) > 0 Then
			; Delete existing first blank child
			_GUICtrlTreeView_Delete($hTreeView, _GUICtrlTreeView_GetFirstChild ($hTreeView, $hDefFolder_Item))
			; Get new first item
			$hDefFolder_Item = _GUICtrlTreeView_GetFirstChild($hTreeView, $hDefFolder_Item)
		EndIf
	Next

	; Clear flag to re-enable expand dummy
	$g_bCFF_AutoExpand = False

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_File_Visible
; Description ...: Ensure files visible if selecting files only and displaying both files and folders
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_File_Visible($hTreeView, $aFill_Ret)

	; Last file visible
	_GUICtrlTreeView_EnsureVisible($hTreeView, $aFill_Ret[2])
	; Last folder visible - first file immediately follows
	_GUICtrlTreeView_EnsureVisible($hTreeView, $aFill_Ret[1])

EndFunc   ;==>_CFF_File_Visible

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Check_Display
; Description ...: Checks for valid Display parameter
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_Check_Display($iDisplay, $sFile_Mask, ByRef $bShow_Ext, ByRef $iHide_HS, ByRef $bBoth_Selectable, ByRef $bDuplicates_Allowed, ByRef $bFileScroll, ByRef $bSplash, ByRef $bDeepest)

	; Show hidden and system files
	If BitAND($iDisplay, 4) Then
		$iHide_HS += 4
		$iDisplay -= 4
	EndIf
	If BitAND($iDisplay, 8) Then
		$iHide_HS += 8
		$iDisplay -= 8
	EndIf
	; Both files and folders selectable
	If BitAND($iDisplay, 16) Then
		$bBoth_Selectable = True
		$iDisplay -= 16
	EndIf
	; Allow duplicate selections
	If BitAND($iDisplay, 32) Then
		$bDuplicates_Allowed = True
		$iDisplay -= 32
	EndIf
	; Hide file extensions?
	If BitAND($iDisplay, 64) Then
		$iDisplay -= 64
		; Check that only one ext is specified
		StringReplace($sFile_Mask, ";", "")
		Local $iExt = @extended
		If $sFile_Mask <> "*.*" And $iExt = 0 Then
			; File exts hidden
			$bShow_Ext = False
		EndIf
	EndIf
	; Scroll to first file
	If BitAND($iDisplay, 128) Then
		$bFileScroll = True
		$iDisplay -= 128
	EndIf
	; Display splashscreens
	If BitAND($iDisplay, 256) Then
		$bSplash = True
		$iDisplay -= 256
	EndIf
	; Only return deepest checked
	If BitAND($iDisplay, 512) Then
		$bDeepest = True
		$iDisplay -= 512
	EndIf
	; Check valid parameter
	Switch $iDisplay
		Case 0 To 3
			Return $iDisplay
		Case Else
			Return SetError(1, 0, 0)
	EndSwitch

EndFunc   ;==>_CFF_Check_Display

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Verify_PreCheck
; Description ...: Checks for valid pre-check array elements
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_Verify_PreCheck(ByRef $g_aCFF_PreCheck, $iDisplay, $bBoth_Selectable)

	Local $sPath, $sAttribs

	For $i = 0 To UBound($g_aCFF_PreCheck) - 1
		; Extract item
		$sPath = $g_aCFF_PreCheck[$i]
		; Get Item sttribs
		$sAttribs = FileGetAttrib($sPath)
		If @error Then
			; Invalid path
			$sPath = ""
			; No point in continuing
			ContinueLoop
		EndIf

		Switch $iDisplay
			Case 0 ; Files only  - unless both selectable flag set
				If StringInStr($sAttribs, "D") Then
					If $bBoth_Selectable Then
						; Add a trailing \ if required
						If StringRight($sPath, 1) <> "\" Then $sPath &= "\"
					Else
						; Remove from list
						$sPath = ""
					EndIf
				EndIf
			Case 1 ; Files only
				If StringInStr($sAttribs, "D") Then
					; Remove from list
					$sPath = ""
				EndIf
			Case 2, 3 ; Folders only
				If StringInStr($sAttribs, "D") Then
					; Add a trailing \ if required
					If StringRight($sPath, 1) <> "\" Then $sPath &= "\"
				Else
					; Remove from list
					$sPath = ""
				EndIf
		EndSwitch
	Next

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Check_Valid
; Description ...: Checks selection is valid
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_Check_Selection($cTV, $hItem, $sRoot, $iDisplay, $bBoth_Selectable, $bShow_Ext, $sFile_Mask)

	; Get full path
	Local $sSelectedPath = $sRoot & StringReplace(_GUICtrlTreeView_GetTree($cTV, $hItem), "|", "\")
	Switch $iDisplay
		Case 0 ; Files and folders displayed
			; Check if folder by looking at attributes
			StringReplace(FileGetAttrib($sSelectedPath), "D", "")
			If @extended Then
				; Are folders selectable?
				If $bBoth_Selectable Then
					; Add trailing \
					$sSelectedPath &= "\"
				Else
					; Reject selection
					$sSelectedPath = ""
				EndIf
			Else
				; File so check extension display
				Continuecase
			EndIf
		Case 1 ; Only files displayed
			; Hide extension?
			If $bShow_Ext = False Then
				$sSelectedPath &= StringTrimLeft($sFile_Mask, 1)
			EndIf
		;Case 2, 3 ; Only folders displayed
	EndSwitch

	Return $sSelectedPath

EndFunc   ;==>_CFF_Check_Selection

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_ParseTV
; Description ...: Gets current TV checkbox state
; Author ........: Melba23
; ===============================================================================================================================
Func _CFF_ParseTV($hTV)

	; Basic check data array and item count
	Local $aParseTV[10][3], $iParseCount = 0

	; Work through TreeView items
	Local $hHandle = _GUICtrlTreeView_GetFirstItem($hTV)
	While 1
		; Add item to array
		$aParseTV[$iParseCount][0] = $hHandle
		$aParseTV[$iParseCount][1] = _GUICtrlTreeView_GetChecked($hTV, $hHandle)
		$aParseTV[$iParseCount][2] = _GUICtrlTreeView_GetText($hTV, $hHandle)
		; Increase count
		$iParseCount += 1
		; Enlarge array if required (minimizes ReDim usage)
		If $iParseCount > UBound($aParseTV) - 1 Then
			ReDim $aParseTV[$iParseCount * 2][3]
		EndIf
		; Move to next item
		$hHandle = _GUICtrlTreeView_GetNext($hTV, $hHandle)
		; Exit if at end
		If $hHandle = 0 Then ExitLoop
	WEnd
	; Remove any empty array elements
	ReDim $aParseTV[$iParseCount][3]

	Return $aParseTV

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Adjust_Parents
; Description ...: Adjusts checkboxes above the one changed
; Author ........: Melba23
; ===============================================================================================================================
Func _CFF_Adjust_Parents($hTV, $hPassedItem, ByRef $aTVCheckData, $bState = True)

	; Get handle of parent
	Local $hParent = _GUICtrlTreeView_GetParentHandle($hTV, $hPassedItem)
	If $hParent = 0 Then Return
	; Assume parent is to be adjusted
	Local $bAdjustParent = True
	; Find parent in array
	Local $iItemIndex = _ArraySearch($aTVCheckData, $hParent)
	; Need to confirm all siblings clear before clearing parent
	If $bState = False Then
		; Check on number of siblings
		Local $iCount = _GUICtrlTreeView_GetChildCount($hTV, $hParent)
		; If only 1 sibling then parent can be cleared - if more then need to look at them all
		If $iCount <> 1 Then
			; Number of siblings checked
			Local $iCheckCount = 0
			; Move through previous siblings
			Local $hSibling = $hPassedItem
			While 1
				$hSibling = _GUICtrlTreeView_GetPrevSibling($hTV, $hSibling)
				; If found
				If $hSibling Then
					; Is sibling checked)
					If _GUICtrlTreeView_GetChecked($hTV, $hSibling) Then
						; Increase count if so
						$iCheckCount += 1
					EndIf
				Else
					; No point in continuing
					ExitLoop
				EndIf
			WEnd
			; Move through later siblings
			$hSibling = $hPassedItem
			While 1
				$hSibling = _GUICtrlTreeView_GetNextSibling($hTV, $hSibling)
				If $hSibling Then
					If _GUICtrlTreeView_GetChecked($hTV, $hSibling) Then
						$iCheckCount += 1
					EndIf
				Else
					ExitLoop
				EndIf
			WEnd
			; If at least one sibling checked then do not clear parent
			If $iCheckCount Then $bAdjustParent = False
		EndIf
	EndIf
	; If parent is to be adjusted
	If $bAdjustParent Then
		; Adjust the array
		$aTVCheckData[$iItemIndex][1] = $bState
		; Adjust the parent
		_GUICtrlTreeView_SetChecked($hTV, $hParent, $bState)
		; And now do the same for the generation above
		_CFF_Adjust_Parents($hTV, $hParent, $aTVCheckData, $bState)
	EndIf

EndFunc   ;==>__GTVEx_Adjust_Parents

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_Adjust_Children
; Description ...: Adjusts checkboxes below the one changed
; Author ........: Melba23
; ===============================================================================================================================
Func _CFF_Adjust_Children($hTV, $hPassedItem, ByRef $aTVCheckData, $bState = True)

	Local $iItemIndex

	; Get the handle of the first child
	Local $hChild = _GUICtrlTreeView_GetFirstChild($hTV, $hPassedItem)
	If $hChild = 0 Then Return
	While 1
		; Find child index
		$iItemIndex = _ArraySearch($aTVCheckData, $hChild)
		If Not @error Then
			; Adjust the array
			$aTVCheckData[$iItemIndex][1] = $bState
			; Adjust the child
			_GUICtrlTreeView_SetChecked($hTV, $hChild, $bState)
			; And now do the same for the generation beow
			_CFF_Adjust_Children($hTV, $hChild, $aTVCheckData, $bState)
			; Now get next child
			$hChild = _GUICtrlTreeView_GetNextChild($hTV, $hChild)
			; Exit the loop if no more found
			If $hChild = 0 Then ExitLoop
		EndIf
	WEnd

EndFunc   ;==>_CFF_Adjust_Children

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_List_Add
; Description ...: Adds item to return list
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_List_Add($sAddFile_List, $sSelectedPath, $cList, $bDuplicates_Allowed, $cTreeView)

	; Check for existing instance in list
	If StringInStr($sAddFile_List, $sSelectedPath & "|") Then
		; If duplicates are allowed
		If $bDuplicates_Allowed Then
			; Add to return string
			$sAddFile_List &= $sSelectedPath & "|"
			; Add to onscreen list
			GUICtrlSendMsg($cList, 0x0180, 0, $sSelectedPath) ; $LB_ADDSTRING
		EndIf
	Else
		; Add to return string
		$sAddFile_List &= $sSelectedPath & "|"
		; Add to onscreen list
		GUICtrlSendMsg($cList, 0x0180, 0, $sSelectedPath) ; $LB_ADDSTRING
	EndIf
	; Scroll to bottom of list
	GUICtrlSendMsg($cList, 0x197, GUICtrlSendMsg($cList, 0x18B, 0, 0) - 1, 0) ; $LB_SETTOPINDEX, $LB_GETCOUNT
	; Return focus to TV
	GUICtrlSetState($cTreeView, 256) ; $GUI_FOCUS

	Return $sAddFile_List

EndFunc   ;==>_CFF_List_Add

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _CFF_List_Del
; Description ...: Deletes item from return list
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func _CFF_List_Del($sAddFile_List, $sSelectedPath, $cList, $cTreeView)

	Local $iTopIndex = 9999

	; Check for existing instance in list
	If StringInStr($sAddFile_List, $sSelectedPath & "|") Then
		; Delete last instance of item
		$sAddFile_List = StringReplace($sAddFile_List, $sSelectedPath & "|", "", -1)
		; Get current scroll position if list selection
		If Not $g_bCFF_ActiveTV Then
			; Get current top index
			$iTopIndex = GUICtrlSendMsg($cList, 0x018E, 0, 0) ; $LB_GETTOPINDEX
		EndIf
		; Replace list content with new list
		GUICtrlSetData($cList, "|" & $sAddFile_List)
	EndIf
	; Scroll as required
	If $iTopIndex <> 9999 Then
		; Scroll to same place if list selection
		GUICtrlSendMsg($cList, 0x197, $iTopIndex, 0) ; $LB_SETTOPINDEX
	Else
		; Scroll to bottom of list if treeview selection
		GUICtrlSendMsg($cList, 0x197, GUICtrlSendMsg($cList, 0x18B, 0, 0) - 1, 0) ; $LB_SETTOPINDEX, $LB_GETCOUNT
	EndIf
	; Return focus to TV
	GUICtrlSetState($cTreeView, 256) ; $GUI_FOCUS

	Return $sAddFile_List

EndFunc   ;==>_CFF_List_Del