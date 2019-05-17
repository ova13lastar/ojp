; #INDEX# =======================================================================================================================
; Title .........: OJP
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
#AutoIt3Wrapper_Res_ProductName=OJP
#AutoIt3Wrapper_Res_Description=Permet d'enregistrer la liasse d'ouverture d'une journée PROGRES en pdf sur un serveur partagé
#AutoIt3Wrapper_Res_ProductVersion=1.0.4
#AutoIt3Wrapper_Res_FileVersion=1.0.4
#AutoIt3Wrapper_Res_CompanyName=CNAMTS/CPAM_ARTOIS/APPLINAT
#AutoIt3Wrapper_Res_LegalCopyright=yann.daniel@assurance-maladie.fr
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Compatibility=Win7
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Icon="static\icon.ico"
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_Run_Au3Stripper=N
#Au3Stripper_Parameters=/MO /RSLN
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; Includes YD
#include "D:\Autoit_dev\Include\YDGVars.au3"
#include "D:\Autoit_dev\Include\YDLogger.au3"
#include "D:\Autoit_dev\Include\YDTool.au3"
; Includes Constants
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
; Includes
#include <String.au3>
; Options
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
OnAutoItExitRegister("_YDTool_ExitApp")
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sAppName", _YDTool_GetAppWrapperRes("ProductName"))
_YDGVars_Set("sAppDesc", _YDTool_GetAppWrapperRes("Description"))
_YDGVars_Set("sAppVersion", _YDTool_GetAppWrapperRes("FileVersion"))
_YDGVars_Set("sAppContact", _YDTool_GetAppWrapperRes("LegalCopyright"))
_YDGVars_Set("sAppVersionV", "v" & _YDGVars_Get("sAppVersion"))
_YDGVars_Set("sAppTitle", _YDGVars_Get("sAppName") & " - " & _YDGVars_Get("sAppVersionV"))
_YDGVars_Set("sAppDirDataPath", @ScriptDir & "\data")
_YDGVars_Set("sAppDirStaticPath", @ScriptDir & "\static")
_YDGVars_Set("sAppDirLogsPath", @ScriptDir & "\logs")
_YDGVars_Set("sAppDirVendorPath", @ScriptDir & "\vendor")
_YDGVars_Set("sAppIconPath", @ScriptDir & "\static\icon.ico")
_YDGVars_Set("sAppConfFile", @ScriptDir & "\conf.ini")
_YDGVars_Set("iAppNbDaysToKeepLogFiles", 15)

_YDGVars_Set("sAppProgresTitle", "PROGRES")
_YDGVars_Set("sAppProgresOuvertureTitle", "PROGRES: Date d'ouverture de la Journée PROGRES")
_YDGVars_Set("sAppProgresExeName", "LNPG.exe")
_YDGVars_Set("sAppProgresExeFile", "C:\PROGRES\" & _YDGVars_Get("sAppProgresExeName"))
_YDGVars_Set("sAppProgresTechLogPath", "C:\PROGRES\LOG")
_YDGVars_Set("sAppProgresTechLogFile", _YDGVars_Get("sAppProgresTechLogPath") & "\TECH_" & @YEAR & @MON & @MDAY & ".LOG")
_YDGVars_Set("sAppProgresNticLogFile", _YDGVars_Get("sAppProgresTechLogPath") & "\NTIC_" & @YEAR & @MON & @MDAY & ".LOG")

_YDLogger_Init()
_YDLogger_LogAllGVars()
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
If Not _YDTool_IsSingleton() Then Exit
;------------------------------
; On supprime les anciens fichiers de log
_YDTool_DeleteOldFiles(_YDGVars_Get("sAppDirLogsPath"), _YDGVars_Get("iAppNbDaysToKeepLogFiles"))
;------------------------------
; On gere l'affichage de l'icone dans le tray
TraySetIcon(_YDGVars_Get("sAppIconPath"))
TraySetToolTip(_YDGVars_Get("sAppTitle"))
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
;------------------------------
Global $g_hWndActive, $g_iPIDActive, $g_hGUI, $g_LblPrinter
Global $g_iLastLineTechLogFile    = 1
Global $g_iLastLineNticLogFile	  = 1
;------------------------------
; On recupere les valeurs de conf.ini
Global $g_sPdfCreatorPrinter = _YDTool_GetAppConfValue("general", "printer")
Global $g_sProgresOuvertureArrasPath = _YDTool_GetAppConfValue("progres", "progres_ouverture_arras_path")
Global $g_sProgresOuvertureLensPath = _YDTool_GetAppConfValue("progres", "progres_ouverture_lens_path")
;------------------------------
; On recupere d autres variables globales
Global $g_sDefaultPrinter = _YDTool_GetDefaultPrinter(@ComputerName)
_YDLogger_Var("$g_sDefaultPrinter", $g_sDefaultPrinter)
Global $g_sDefaultPrinterName = StringReplace($g_sDefaultPrinter, "\\w11620101rps001\", "")
_YDLogger_Var("$g_sDefaultPrinterName", $g_sDefaultPrinterName)
Global $g_sSite = _YDTool_GetHostSite(@ComputerName)
_YDLogger_Var("$g_sSite", $g_sSite)
Global $g_sSiteNetworkPath = ($g_sSite = "ARRAS") ? $g_sProgresOuvertureArrasPath : $g_sProgresOuvertureLensPath
_YDLogger_Var("$g_sSiteNetworkPath", $g_sSiteNetworkPath)
; #MAIN SCRIPT# =================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
	Global $iMsg = TrayGetMsg()
	Select
		Case $iMsg = $idTrayExit
			_YDTool_ExitConfirm()
		Case $iMsg = $idTrayAbout
			_YDTool_GUIShowAbout()
		Case Else
			_Main()
	EndSelect
	;------------------------------
    Sleep(10)
WEnd
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement principal
; Syntax ........: _Main()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 15/05/2019
; Notes .........:
;================================================================================================================================
Func _Main()
	Local $sFuncName = "_Main"
	; On recupere le Handle de la fenetre active et son PID
	$g_hWndActive = WinGetHandle("[ACTIVE]")
	$g_iPIDActive = WinGetProcess($g_hWndActive)

	; On ne travaille que si PROGRES est lance et si la fenetre d'ouverture de la journée PROGRES est détectée
	If ProcessExists(_YDGVars_Get("sAppProgresExeName")) And WinActive(_YDGVars_Get("sAppProgresOuvertureTitle")) Then
		_YDLogger_Log("Fenetre ouverture PROGRES : ouverte", $sFuncName)
		; On recupere l'UGE
		Local $sTechUge = _GetUGEFromTechLogFile()
		; On modifie le registre pour modifier le Path et le nom du fichier
		_YDLogger_Log("Modification du registre", $sFuncName, 2)
		RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveDirectory", "REG_SZ", $g_sSiteNetworkPath & "\")
		RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveFilename", "REG_SZ", "<Datetime>_" & $sTechUge)
		; On bascule sur PDFCreator
		_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
		While _YDTool_GetDefaultPrinter(@ComputerName) <> $g_sPdfCreatorPrinter
			Sleep(100)
		WEnd
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
		; On attend que la fenetre d'ouvertture progres soit fermee
		WinWaitClose(_YDGVars_Get("sAppProgresOuvertureTitle"))
		_YDLogger_Log("Fenetre ouverture PROGRES : fermee", $sFuncName)
		; Tant que PROGRES existe, on verifie que l'UGE du NTIC_xxxx.LOG soit egale à la nouvelle UGE
		While ProcessExists(_YDGVars_Get("sAppProgresExeName"))
			Local $sNticUge = _GetUGEFromNticLogFile()
			If $sNticUge = $sTechUge Then
				_YDLogger_Log("Bascule sur nouvelle UGE OK ! ", $sFuncName)
				ExitLoop
			EndIf
			Sleep(100)
		WEnd
		; On retourne sur l'imprimante par defaut
		_YDTool_SetDefaultPrinter($g_sDefaultPrinter)
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Retour sur imprimante : " & $g_sDefaultPrinterName, 5000)
		; On propose d'ouvrir le dossier
		If ProcessExists(_YDGVars_Get("sAppProgresExeName")) Then
			Local $iInputOpenFolder = MsgBox(4, _YDGVars_Get("sAppTitle"), "Souhaitez-vous ouvrir le dossier des liasses d'Ouverture de Journée PROGRES ?")
			If ($iInputOpenFolder = 6) Then
				ShellExecute($g_sSiteNetworkPath)
			Endif
		EndIf
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: permet de récupérer l'UGE via le fichier TECH_xxxxxxx.LOG
; Syntax ........: _GetUGEFromTechLogFile()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 15/05/2019
; Notes .........:
;================================================================================================================================
Func _GetUGEFromTechLogFile()
	Local $sFuncName = "_GetUGEFromTechLogFile"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "F_LNPG_USER0 Appel du centre : "
	Local $sUGE = "0000"
	; On ne fait des recherches que si la fenetre d'ouverture de la journée PROGRES est détectée
	If WinExists(_YDGVars_Get("sAppProgresOuvertureTitle")) Then
		$hLogFile = FileOpen(_YDGVars_Get("sAppProgresTechLogFile"), 0)
		If $hLogFile = -1 Then
			_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
			_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier de log inaccessible !", 0, $TIP_ICONASTERISK)
			Return False
		Endif
		; On recupere le nombre de lignes du fichier de log
		Local $iFileCountLine = _FileCountLines(_YDGVars_Get("sAppProgresTechLogFile"))
		_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
		; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
		If $iFileCountLine < $g_iLastLineTechLogFile Then $g_iLastLineTechLogFile = 1
		_YDLogger_Var("$g_iLastLineTechLogFile (avant)", $g_iLastLineTechLogFile, $sFuncName, 2)
		; On boucle sur le fichier log
		For $i = $iFileCountLine to $g_iLastLineTechLogFile Step -1
			$sFileLine = FileReadLine($hLogFile, $i)
			Local $iUGEDetected = StringInStr($sFileLine, $sPattern, 0, 1, 1)
			; Si pattern trouve on sort de la boucle
			If $iUGEDetected > 0 Then
				;------------------------------
				_YDLogger_Log("Pattern trouve : " & $sPattern, $sFuncName, 2)
				$sUGE = StringRight($sFileLine, 4)
				$g_iLastLineTechLogFile = _FileCountLines(_YDGVars_Get("sAppProgresTechLogFile")) + 1
				ExitLoop
			EndIf
		Next
		FileClose($hLogFile)
		;------------------------------
		; On log les infos utiles
		_YDLogger_Var("$sUGE", $sUGE, $sFuncName)
		_YDLogger_Var("$g_iLastLineTechLogFile (apres)", $g_iLastLineTechLogFile, $sFuncName, 2)
		;------------------------------
		Return $sUGE
	Endif
	Return $sUGE
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: permet de récupérer l'UGE via le fichier NTIC_xxxxxxx.LOG
; Syntax ........: _GetUGEFromNticLogFile()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 15/05/2019
; Notes .........:
;================================================================================================================================
Func _GetUGEFromNticLogFile()
	Local $sFuncName = "_GetUGEFromNticLogFile"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "Sz_UgeNum : "
	Local $sUGE = "0000"
	; On ne fait des recherches que si PROGRES est lance
	If ProcessExists(_YDGVars_Get("sAppProgresExeName")) Then
		If FileExists(_YDGVars_Get("sAppProgresNticLogFile")) = 0 Then
			_YDLogger_Log("Fichier non present : " & _YDGVars_Get("sAppProgresNticLogFile"), $sFuncName)
			Return False
		EndIf
		$hLogFile = FileOpen(_YDGVars_Get("sAppProgresNticLogFile"), 0)
		If $hLogFile = -1 Then
			_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
			_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier de log inaccessible !", 0, $TIP_ICONASTERISK)
			Return False
		Endif
		; On recupere le nombre de lignes du fichier de log
		Local $iFileCountLine = _FileCountLines(_YDGVars_Get("sAppProgresNticLogFile"))
		_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
		; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
		If $iFileCountLine < $g_iLastLineNticLogFile Then $g_iLastLineNticLogFile = 1
		_YDLogger_Var("$g_iLastLineNticLogFile (avant)", $g_iLastLineNticLogFile, $sFuncName, 2)
		; On boucle sur le fichier log
		For $i = $iFileCountLine to $g_iLastLineNticLogFile Step -1
			$sFileLine = FileReadLine($hLogFile, $i)
			Local $iUGEDetected = StringInStr($sFileLine, $sPattern, 0, 1, 1)
			; Si pattern trouve on sort de la boucle
			If $iUGEDetected > 0 Then
				;------------------------------
				_YDLogger_Log("Pattern trouve : " & $sPattern, $sFuncName, 2)
				$sUGE = StringRight($sFileLine, 4)
				$g_iLastLineNticLogFile = _FileCountLines(_YDGVars_Get("sAppProgresNticLogFile")) + 1
				ExitLoop
			EndIf
		Next
		FileClose($hLogFile)
		;------------------------------
		; On log les infos utiles
		_YDLogger_Var("$sUGE", $sUGE, $sFuncName)
		_YDLogger_Var("$g_iLastLineTechLogFile (apres)", $g_iLastLineNticLogFile, $sFuncName, 2)
		;------------------------------
		Return $sUGE
	Endif
	Return $sUGE
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de vérifier si le pdf a été imprimé
; Syntax.........: _IsPrinted($g_sSiteNetworkPath, $sUGE)
; Parameters ....: $g_sSiteNetworkPath - Chemin reseau de destination
;				   $sUGE 			 - UGE traitee
; Return values .: Success      - True
;                  Failure      - False
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/05/2019
; Notes .........:
;================================================================================================================================
;~ Func _IsPrinted($sUGE)
;~ 	Local $sFuncName = "_IsPrinted"
;~ 	_YDLogger_Var("$sUGE", $sUGE, $sFuncName)
;~ 	Local $aFiles = _FileListToArrayEx($g_sSiteNetworkPath & "\", @YEAR & @MON & @MDAY  & "*" & "_" & $sUGE & ".pdf", 1, 0, 0)
;~ 	If @error Then Return False
;~ 	If $aFiles[0] > 0 Then
;~ 		Return True
;~ 	EndIf
;~ 	Return False
;~ EndFunc



