
#include "ChooseFileFolder.au3"

Global $sRet
Global $sRootFolder = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", Default, Default))
If StringRight($sRootFolder, 5) == "beta\" Then
    $sRootFolder = StringTrimRight($sRootFolder, 5)
EndIf
ConsoleWrite($sRootFolder & @CRLF)

; Register WM_NOTIFY handler
$sRet = _CFF_RegMsg()
If Not $sRet Then
	MsgBox(16, "Failure!", "Handler not registered")
	Exit
EndIf

; All ready drives - no default
_CFF_Choose("Ex 1a: All drives with no default", 300, 500)
ConsoleWrite("Ex 1a: " & @error & @CRLF)

_CFF_Choose("Ex 1b: All drives with no default (combo)", 300, 500, Default, Default, "||c")
ConsoleWrite("Ex 1b: " & @error & @CRLF)


; Limit to certain drives - note only ready drives are displayed in the combo
_CFF_Choose("Ex 2: Drive list with no default", 300, 500, Default, Default, "acmnxyz")
ConsoleWrite("Ex 2: " & @error & @CRLF)

; Open the default drive of a set
_CFF_Choose("Ex 3: Drive list with default", 300, 500, Default, Default, "cmn|n")
ConsoleWrite("Ex 3: " & @error & @CRLF)

; Open the default drive and show all drives in combo
_CFF_Choose("Ex 4a: All drives with default drive", 300, 500, Default, Default, "|n")
ConsoleWrite("Ex 4a: " & @error & @CRLF)

_CFF_Choose("Ex 4b: All drives with default folder", 300, 500, Default, Default, "|" & $sRootFolder)
ConsoleWrite("Ex 4b: " & @error & @CRLF)

; This should fail with error 1 - Path does not exist or invalid drive list
_CFF_Choose("Ex 5", 300, 500, Default, Default, "C:\Fr:ed\") ; Invalid path
ConsoleWrite("Ex 5: " & @error & @CRLF)

; This should fail with error 1 - Path does not exist or invalid drive list
_CFF_Choose("Ex 6", 300, 500, Default, Default, "cde5") ; Digit in drive list
ConsoleWrite("Ex 6: " & @error & @CRLF)

; Multiple default drives uses the first
_CFF_Choose("Ex 7", 300, 500, Default, Default, "|cde")
ConsoleWrite("Ex 7: " & @error & @CRLF)

; The default drive not in drive list and so is ignored
_CFF_Choose("Ex 8: Default not in drive list", 300, 500, Default, Default, "cde|f")
ConsoleWrite("Ex 8: " & @error & @CRLF)

; The chosen default drive does not exist and so is ignored
_CFF_Choose("Ex 9: Default drive not ready", 300, 500, Default, Default, "cdez|z")
ConsoleWrite("Ex 9: " & @error & @CRLF)