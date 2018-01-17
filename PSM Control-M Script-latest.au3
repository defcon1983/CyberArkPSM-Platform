
;============================================================
;             PSM AutoIt Dispatcher Skeleton
;             ------------------------------
;
; Use this skeleton to create your own
; connection components integrated with the PSM.
; Areas you may want to modify are marked
; Created by: Muthaheer Khan
; June 2016
; Please note that this platform is created based my experiance in AutoIT, 
;============================================================


#AutoIt3Wrapper_UseX64=n
Opt("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 3) ; EXACT_MATCH!
AutoItSetOption("WinDetectHiddenText",1)

#include "PSMGenericClientWrapper.au3"

;=======================================
; Consts & Globals
;=======================================
Global Const $DISPATCHER_NAME									= "BMC Control-M" ; CHANGE_ME
Global Const $CLIENT_EXECUTABLE									= "C:\Program Files\BMC Software\Control-M EM 8.0.00\Default\bin\emwa.exe"
Global Const $ERROR_MESSAGE_TITLE  								= "PSM " & $DISPATCHER_NAME & " Dispatcher error message"
Global Const $LOG_MESSAGE_PREFIX 								= $DISPATCHER_NAME & " Dispatcher - "
Global Const $APPLICATION_TITLE 								= "Control-M Workload Automation"

Global $TargetUsername
Global $TargetPassword
Global $TargetAddress
Global $TargetLogonDomain
Global $ExecutableParameters
Global $ExecutableWithParameters
Global $ConnectionClientPID		= 0

;Control-M remote server address
Global $TargetServer			= "rfintctm"
Global $TargetHost				= "rfintctm"



;=======================================
; Code
;=======================================
Exit Main()

;=======================================
; Main
;=======================================
Func Main()

	; Init PSM Dispatcher utils wrapper
	ToolTip ("Initializing...")
	if (PSMGenericClient_Init() <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	LogWrite("successfully initialized Dispatcher Utils Wrapper")

	; Get the dispatcher parameters
	FetchSessionProperties()

	LogWrite("mapping local drives")
	if (PSMGenericClient_MapTSDrives() <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	LogWrite("starting client application")
	ToolTip ("Starting " & $DISPATCHER_NAME & "...")


	; -------------------------------------------------------------------------------------------------------------------
	; 												Handle login here!
	; -------------------------------------------------------------------------------------------------------------------

    ;Adding parameters to the executable
	$ExecutableParameters = " -u " & $TargetUsername & " -p " & $TargetPassword & " -s " & $TargetServer & " --nshost " & $TargetHost
	$ExecutableWithParameters = $CLIENT_EXECUTABLE & $ExecutableParameters

	;Run the executable with the parameters
	$ConnectionClientPID = Run($ExecutableWithParameters)
;	$ConnectionClientPID = Run($CLIENT_EXECUTABLE)

	if ($ConnectionClientPID == 0) Then
		Error(StringFormat("Failed to execute process [%s]", $CLIENT_EXECUTABLE, @error))
	EndIf

	; Send PID to PSM as early as possible so recording/monitoring can begin
	LogWrite("sending PID to PSM")
	if (PSMGenericClient_SendPID($ConnectionClientPID) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf


	; Terminate PSM Dispatcher utils wrapper
	LogWrite("Terminating Dispatcher Utils Wrapper")
	PSMGenericClient_Term()
	Return $PSM_ERROR_SUCCESS
EndFunc

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
	if (PSMGenericClient_IsInitialized()) Then
		LogWrite($ErrorMessage, True)
		PSMGenericClient_Term()
	EndIf

	Local $MessageFlags = BitOr(0, 16, 262144) ; 0=OK button, 16=Stop-sign icon, 262144=MsgBox has top-most attribute set

	MsgBox($MessageFlags, $ERROR_MESSAGE_TITLE, $ErrorMessage)

	; If the connection component was already invoked, terminate it
	if ($ConnectionClientPID <> 0) Then
		ProcessClose($ConnectionClientPID)
		$ConnectionClientPID = 0
	EndIf

	Exit $Code
EndFunc


Func LogWrite($sMessage, $LogLevel = $LOG_LEVEL_TRACE)
	Return PSMGenericClient_LogWrite($LOG_MESSAGE_PREFIX & $sMessage, $LogLevel)
EndFunc


Func FetchSessionProperties() ; CHANGE_ME
	if (PSMGenericClient_GetSessionProperty("Username", $TargetUsername) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	if (PSMGenericClient_GetSessionProperty("Password", $TargetPassword) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	if (PSMGenericClient_GetSessionProperty("Address", $TargetAddress) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

EndFunc