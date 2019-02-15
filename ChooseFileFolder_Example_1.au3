
#include <MsgBoxConstants.au3>

#include "ChooseFileFolder.au3"

Local $sRet, $aRet
Global $sRootFolder = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", Default, -1))
ConsoleWrite($sRootFolder & @CRLF)

; Register handlers
$sRet = _CFF_RegMsg()
If Not $sRet Then
	MsgBox(16, "Failure!", "Handler not registered")
	Exit
EndIf

#cs

MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Example Script 1", "Please read the comments for each Example as it runs" & @CRLF & "to see how the display and return values have been set")

; Pick a single folder from within the AutoIt installation folders
; Select using the "Select" button only - doubleclick expands but does not select item
; Make the final parameter 3 if you want to use doubleclick as well - but beware when clicking on the items!
$sRet = _CFF_Choose("Ex 1: Choose a folder (Button only)", 300, 500, -1, -1, $sRootFolder, "*", 2)
If $sRet Then
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 1", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 1", "No Selection")
EndIf

; Pick a single *.au3 or *.chm file from within the AutoIt installation folders
; All folders are displayed whether they contain files or not
; Use either the "Select" button or a double click - only files can be selected
; Note that passing a filename highlights that file in the tree
$sRet = _CFF_Choose("Ex 2: Choose a file (default highlight)", 300, 500, -1, -1, $sRootFolder & "autoit.chm", "*.au3;*.chm")
If $sRet Then
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 2", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 2", "No Selection")
EndIf

; Pick any file on any drive
; Hidden and system files/folders will be shown - compare with example 4
; Note flag to display splash screen during search - again compare with example 4
$sRet = _CFF_Choose("Ex 3a: Select a file (all drives)", 300, 500, -1, -1, "", "", 256)
If $sRet Then
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 3a", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 3a", "No Selection")
EndIf

; As above but highlighting defined folder within all drives
; Note flag to display splash screen during expansion
$sRet = _CFF_Choose("Ex 3b: Choose a file (default folder highlight)", 300, 400, -1, -1, "|" & $sRootFolder, "", 256)
If $sRet Then
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 3b", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 3b", "No Selection")
EndIf

; Pick multiple files from any drive - list takes 20% of dialog
; Play with the final parameter to change that percentage
; Hidden and system files/folders will not be shown ( + 4 + 8)
; Multiple instances of the same item allowed ( + 32)
; Use either the "Add" button or a double click to add to the list - only files can be added
; If Ctrl key pressed as item selected, last instance of item is removed from list
; Press "Return" button when selection ended to get "|" delimited string of selected files
$sRet = _CFF_Choose("Ex 4: Select multiple files (duplicates allowed)", 300, 500, -1, -1, "", Default, 0 + 4 + 8 + 32, 20)
If $sRet Then
	$aRet = StringSplit($sRet, "|")
	$sRet = ""
	For $i = 1 To $aRet[0]
		$sRet &= $aRet[$i] & @CRLF
	Next
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 4", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 4", "No Selection")
EndIf

; As above but with the ability to choose files AND folders (+ 16)
; Multiple instances not permitted (Not + 32)
; Beware when double clicking on the folders - it expands but also selects them!
; Selected folders have a trailing "\" added
; Note negative width to make dialog resizeable
$sRet = _CFF_Choose("Ex 5: Select multiple files AND folders (resizeable)", -300, 500, -1, -1, "", Default, 0 + 4 + 8 + 16, 20)
If $sRet Then
	$aRet = StringSplit($sRet, "|")
	$sRet = ""
	For $i = 1 To $aRet[0]
		$sRet &= $aRet[$i] & @CRLF
	Next
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 5", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 5", "No Selection")
EndIf

; Choose a single file from a specified folder - no subfolders displayed
; Note file exts displayed because of *.* mask even though requested to be hidden (+ 64)
$sRet = _CFF_Choose("Ex 6: Select a file", 300, 500, -1, -1, $sRootFolder, "*.*", 1 + 64)
If $sRet Then
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 6", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 6", "No Selection")
EndIf

; Choose a single file from a specified folder - no subfolders displayed
; Note file exts hidden, but returned (+ 64)
$sRet = _CFF_Choose("Ex 7: Select a file - no extensions displayed", 300, 500, -1, -1, $sRootFolder & "Include\", "*.au3", 1 + 64)
If $sRet Then
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 7 - But extension returned", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 7", "No Selection")
EndIf

MsgBox($MB_SYSTEMMODAL, "Checkboxes", "The following examples all use checkboxes allowing for multiple selections")

; Choose multiple files using checkboxes - all checked items returned
$sRet = _CFF_Choose("Ex 8 - Multiple files only", 300, 500, -1, -1, "", Default, 0, -1)
If $sRet Then
	$aRet = StringSplit($sRet, "|")
	$sRet = ""
	For $i = 1 To $aRet[0]
		$sRet &= $aRet[$i] & @CRLF
	Next
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 8 - Multiple files only", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 8", "No Selection")
EndIf

#ce

; Choose multiple files/folders using checkboxes - all checked items returned
$sRet = _CFF_Choose("Ex 9 - Multiple files and folders", 300, 500, -1, -1, "", Default, 0 + 16, -1)
If $sRet Then
	$aRet = StringSplit($sRet, "|")
	$sRet = ""
	For $i = 1 To $aRet[0]
		$sRet &= $aRet[$i] & @CRLF
	Next
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 9 - Multiple files and folders", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 9", "No Selection")
EndIf

; Pre-check the AutoIt executable file
Local $aPreCheck_List[] = [@AutoItExe]
; Note that pre-checked file will be returned unless expanded and cleared
_CFF_SetPreCheck($aPreCheck_List)
; Only lowest item on a checked tree will be returned
$sRet = _CFF_Choose("Ex 10 - Preset checkbox", 300, 500, -1, -1, "", Default, 0 + 16 + 512, -1)
If $sRet Then
	$aRet = StringSplit($sRet, "|")
	$sRet = ""
	For $i = 1 To $aRet[0]
		$sRet &= $aRet[$i] & @CRLF
	Next
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 10 - Multiple with pre-checked", "Selected:" & @CRLF & @CRLF & $sRet)
Else
	MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Ex 10", "No Selection")
EndIf