#AutoIt3Wrapper_UseX64=Y
Opt("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 3) ; EXACT_MATCH!
AutoItSetOption("MouseCoordMode", 0) ; Active window mouse coord mode

;============================================================
;            PSM MICROSOFT MANAGEMENT CONSOLE
;             ------------------------------
;

; Created : Muthaheer Khan
; Mod : June 2017
;============================================================
#include "PSMGenericClientWrapper.au3"

;=======================================
; Consts & Globals
;=======================================
Global Const $DISPATCHER_NAME = "MMC-Dispatcher" ; CHANGE_ME
Global Const $CLIENT_EXECUTABLE = "C:\Windows\System32\mmc.exe" ; CHANGE_ME
Global Const $ERROR_MESSAGE_TITLE = "PSM " & $DISPATCHER_NAME & " Dispatcher error message"
Global Const $LOG_MESSAGE_PREFIX = $DISPATCHER_NAME & " Dispatcher - "

Global $TargetUsername
Global $TargetPassword
Global $TargetDomain
Global $ClientChosen
Global $ConnectionClientPID = 0

;=======================================
; Code
;=======================================
Exit Main()

;==================================
; Functions
;==================================
; #FUNCTION# ====================================================================================================================
; Name...........: Error
; Description ...: An exception handler - displays an error message and terminates the dispatcher
; Parameters ....: $ErrorMessage - Error message to display
; 				   $Code 		 - [Optional] Exit error code
; ===============================================================================================================================
Func Error($ErrorMessage, $Code = -1)

	; If the dispatcher utils DLL was already initialized, write an error log message and terminate the wrapper
	If (PSMGenericClient_IsInitialized()) Then
		LogWrite($ErrorMessage, True)
		PSMGenericClient_Term()
	EndIf

	Local $MessageFlags = BitOR(0, 16, 262144) ; 0=OK button, 16=Stop-sign icon, 262144=MsgBox has top-most attribute set

	MsgBox($MessageFlags, $ERROR_MESSAGE_TITLE, $ErrorMessage)

	; If the connection component was already invoked, terminate it
	If ($ConnectionClientPID <> 0) Then
		ProcessClose($ConnectionClientPID)
		$ConnectionClientPID = 0
	EndIf

	Exit $Code
EndFunc   ;==>Error

Func FetchSessionProperties() ; CHANGE_ME
	If (PSMGenericClient_GetSessionProperty("Username", $TargetUsername) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	If (PSMGenericClient_GetSessionProperty("Password", $TargetPassword) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	If (PSMGenericClient_GetSessionProperty("LogonDomain", $TargetDomain) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	If (PSMGenericClient_GetSessionProperty("SelectedConnCompID", $ClientChosen) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf
EndFunc   ;==>FetchSessionProperties


Func LogWrite($sMessage, $LogLevel = $LOG_LEVEL_TRACE)
	Return PSMGenericClient_LogWrite($LOG_MESSAGE_PREFIX & $sMessage, $LogLevel)
EndFunc   ;==>LogWrite

;=======================================
; Main
;=======================================
Func Main()

	; Init PSM Dispatcher utils wrapper
	ToolTip("Initialisation...")
	If (PSMGenericClient_Init() <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	LogWrite("successfully initialized Dispatcher Utils Wrapper")

	; Get the dispatcher parameters
	FetchSessionProperties()


	LogWrite("starting client application")
	ToolTip("Starting MMC...")

	Switch $ClientChosen
		Case "PSM-RDP-ADSIEDIT"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\adsiedit.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-ADSSIT"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\dssite.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-AD_Module_PowerShell"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"' & ' -noexit -command "import-module ActiveDirectory; cd ' & @UserProfileDir & '"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-DHCP"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\dhcpmgmt.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-DNSMGMT"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\dnsmgmt.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-DOMAIN"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\domain.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-DSA"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\dsa.msc"' & ' /domain=bycn.bouygues-construction.com', "", @SW_MAXIMIZE)


		Case "PSM-RDP-ExchangeManagementConsole"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Program Files\Microsoft\Exchange Server\V14\Bin\Exchange Management Console.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-ExchangeManagementConsolePowerShell"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"' & "  -noexit -executionPolicy ByPass -command . 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto", "", @SW_MAXIMIZE)


		Case "PSM-RDP-ExchangePowerShell-ISE"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"', "", @SW_MAXIMIZE)

		 Case "PSM-PowerShell"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"', "", @SW_MAXIMIZE)

		Case "PSM-RDP-GPMC"
 			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, 'C:\Windows\System32\mmc.exe "C:\Windows\System32\gpmc.msc"', "", @SW_MAXIMIZE)
			;Run('"runas.exe /netonly /user:" & $TargetDomain & "\" & $TargetUsername & " cmd /c C:\WINDOWS\System32\mmc.exe C:\WINDOWS\system32\gpmc.msc"', "", @SW_MAXIMIZE)

		Case "PSM-RDP-PKIVIEW"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, $CLIENT_EXECUTABLE & ' "C:\Windows\system32\pkiview.msc"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-PowerShell_ISE"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ISE.exe"', "", @SW_MAXIMIZE)


		Case "PSM-RDP-LyncManagementConsolepowershell"
			$ConnectionClientPID = RunAs($TargetUsername, $TargetDomain, $TargetPassword, 2, '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"' & "  -noexit -command cd $env:UserProfile; Import-Module 'C:\Program Files\Common Files\Microsoft Lync Server 2010\Modules\Lync\Lync.psd1'", "", @SW_MAXIMIZE)

		Case Else
			Error("SSO on " & $ClientChosen & " is not implemented yet!")

	EndSwitch

	If ($ConnectionClientPID == 0) Then
		Error(StringFormat("Failed to execute process [%s]", $CLIENT_EXECUTABLE, @error))
	EndIf

	; Send PID to PSM as early as possible so recording/monitoring can begin
	LogWrite("sending PID to PSM")
	If (PSMGenericClient_SendPID($ConnectionClientPID) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	; ------------------
	; Handle login here! ; CHANGE_ME
	; ------------------

	ToolTip("")

	; Terminate PSM Dispatcher utils wrapper
	LogWrite("Terminating Dispatcher Utils Wrapper")
	PSMGenericClient_Term()

	Return $PSM_ERROR_SUCCESS
EndFunc   ;==>Main
