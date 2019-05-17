#include-once

;#AutoIt3Wrapper_au3check_parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
; #FUNCTION# ====================================================================================================================
; Name...........: _FileListToArrayEx
; Description ...: Lists files and\or folders in a specified path (Similar to using Dir with the /B Switch)
; Syntax.........: _FileListToArrayEx($sPath[, $sFilter = "*"[, $iFlag = 0 [, $iSDir = 0 [, $iFPath = 1]]]])
; Parameters ....: $sPath   - Path to generate filelist for.
;                  $sFilter - Optional the filter to use, default is *. (Multiple filter groups such as "*.png|*.jpg|*.bmp") Search the Autoit3 helpfile for the word "WildCards" For details.
;                  $iFlag   - Optional: specifies whether to return files folders or both
;                  		|$iFlag=0 (Default) Return both files and folders
;                  		|$iFlag=1 Return files only
;                  		|$iFlag=2 Return Folders only
;                  $iSDir   - Optional: specifies whether to return files folders or both
;                  		|$iSDir=0 Do not Search subdirectory (default)
;                  		|$iSDir=1 Search subdirectory
;                  $iFPath  - Optional: specifies whether to return files folders or both
;                  		|$iFPath=0 Do not return full path
;                  		|$iFPath=1 Return full path (Default)
; Return values .: @Error
;						|1 = Path not found or invalid
;                  		|2 = Invalid $sFilter
;                  		|3 = Invalid $iFlag
;                  		|4 = No File(s) Found
; Author ........: SolidSnake <metalgx91 at="" gmail="" dot="" com="">
; Modified.......:
; Remarks .......: The array returned is one-dimensional and is made up as follows:
;                                $array[0] = Number of Files\Folders returned
;                                $array[1] = 1st File\Folder
;                                $array[2] = 2nd File\Folder
;                                $array[3] = 3rd File\Folder
;                                $array[n] = nth File\Folder
; Related .......:
; Link ..........: http://www.autoitscript.com/forum/topic/131277-filelisttoarrayex
; Example .......: Yes
; Note ..........: Special Thanks to Helge and Layer for help with the $iFlag update speed optimization by code65536, pdaughe
;                  Update By DXRW4E
;                  Update By Tlem (Path return for non recurcive search. Add use of \\?\ to bypass Windows limits paths to 260 chars)
;				   Update By Tlem to correct bugs introduced by previous modifications (not working with network path).
; ===============================================================================================================================
Func _FileListToArrayEx($sPath, $sFilter = "*", $iFlag = 0, $iSDir = 0, $iFPath = 1)
	Local $hSearch, $sFile, $sFileList, $sDelim = "|", $sSDirFTMP = $sFilter, $FPath
	$sPath = StringRegExpReplace($sPath, "[\\/]+\z", "") & "\" ; ensure single trailing backslash
	If Not FileExists($sPath) Then Return SetError(1, 1, "")
	If StringRegExp($sFilter, "[\\/:><]|(?s)\A\s*\z") Then Return SetError(2, 2, "")
	If Not ($iFlag = 0 Or $iFlag = 1 Or $iFlag = 2) Then Return SetError(3, 3, "")

	If StringLeft($sPath, 2) <> "\\" Then
		$FPath = $sPath
		$sPath = "\\?\" & $sPath
	Else
		$FPath = $sPath
	EndIf

	If $iFPath = 0 Then
		$FPath = ""
	EndIf

	$hSearch = FileFindFirstFile($sPath & "*")
	If @error Then Return SetError(4, 4, "")
	Local $hWSearch = $hSearch, $hWSTMP = $hSearch, $SearchWD, $sSDirF[3] = [0, StringReplace($sSDirFTMP, "*", ""), "(?i)(" & StringRegExpReplace(StringRegExpReplace(StringRegExpReplace(StringRegExpReplace(StringRegExpReplace(StringRegExpReplace("|" & $sSDirFTMP & "|", '\|\h*\|[\|\h]*', "\|"), '[\^\$\(\)\+\[\]\{\}\,\.\=]', "\\$0"), "\|([^\*])", "\|^$1"), "([^\*])\|", "$1\$\|"), '\*', ".*"), '^\||\|$', "") & ")"]
	While 1
		$sFile = FileFindNextFile($hWSearch)
		If @error Then
			If $hWSearch = $hSearch Then ExitLoop
			FileClose($hWSearch)
			$hWSearch -= 1
			$SearchWD = StringLeft($SearchWD, StringInStr(StringTrimRight($SearchWD, 1), "\", 1, -1))
		ElseIf $iSDir Then
			$sSDirF[0] = @extended
			If ($iFlag + $sSDirF[0] <> 2) Then
				If $sSDirF[1] Then
					If StringRegExp($sFile, $sSDirF[2]) Then $sFileList &= $sDelim & $FPath & $SearchWD & $sFile
				Else
					$sFileList &= $sDelim & $FPath & $SearchWD & $sFile
				EndIf
			EndIf
			If Not $sSDirF[0] Then ContinueLoop
			$hWSTMP = FileFindFirstFile($sPath & $SearchWD & $sFile & "\*")
			If $hWSTMP = -1 Then ContinueLoop
			$hWSearch = $hWSTMP
			$SearchWD &= $sFile & "\"
		Else
			If ($iFlag + @extended = 2) Or StringRegExp($sFile, $sSDirF[2]) = 0 Then ContinueLoop
			$sFileList &= $sDelim & $FPath & $sFile
		EndIf
	WEnd
	FileClose($hSearch)
	If Not $sFileList Then Return SetError(4, 4, "")

	Return StringSplit(StringTrimLeft($sFileList, 1), "|")
EndFunc   ;==>_FileListToArrayEx
