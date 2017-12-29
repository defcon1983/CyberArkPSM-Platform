;============================================================;             PSM AutoIt Dispatcher Skeleton;             ------------------------------;; Use this skeleton to create your own; connection components integrated with the PSM.; Areas you may want to modify are marked; Created by: Muthaheer Khan; April 2017; Please note that this platform is created based my experiance in AutoIT, ;============================================================
#AutoIt3Wrapper_UseX64=nOpt("MustDeclareVars", 1)AutoItSetOption("WinTitleMatchMode", 3);#include "PSMGenericClientWrapper.au3"
;=======================================; Consts & Globals;=======================================Global Const $DISPATCHER_NAME            = "vSphere Web Client" ; CHANGE_MEGlobal Const $CLIENT_EXECUTABLE         = "c:\Program Files (x86)\Internet Explorer\iexplore.exe https://vconsole:9443/vsphere-client/#/" ; CHANGE_MEGlobal Const $ERROR_MESSAGE_TITLE          = "PSM " & $DISPATCHER_NAME & " Dispatcher error message"Global Const $LOG_MESSAGE_PREFIX         = $DISPATCHER_NAME & " Dispatcher - "
Global $TargetUsernameGlobal $TargetPasswordGlobal $TargetAddressGlobal $ConnectionClientPID = 0
;=======================================; Code;=======================================Exit Main()
;=======================================; Main;=======================================Func Main()
 ; Muthaheer: Init PSM Dispatcher utils wrapper ToolTip ("Initializing...") if (PSMGenericClient_Init() <> $PSM_ERROR_SUCCESS) Then  Error(PSMGenericClient_PSMGetLastErrorString()) EndIf
 LogWrite("successfully initialized Dispatcher Utils Wrapper")
 ; Muthaheer:Get the dispatcher parameters FetchSessionProperties()
 LogWrite("mapping local drives") if (PSMGenericClient_MapTSDrives() <> $PSM_ERROR_SUCCESS) Then  Error(PSMGenericClient_PSMGetLastErrorString()) EndIf
 LogWrite("starting client application") ToolTip ("Starting " & $DISPATCHER_NAME & "...") $ConnectionClientPID = Run($CLIENT_EXECUTABLE) if ($ConnectionClientPID == 0) Then  Error(StringFormat("Failed to execute process [%s]", $CLIENT_EXECUTABLE, @error)) EndIf
 ;  Muthaheer: Send PID to PSM as early as possible so recording/monitoring can begin LogWrite("sending PID to PSM") if (PSMGenericClient_SendPID($ConnectionClientPID) <> $PSM_ERROR_SUCCESS) Then  Error(PSMGenericClient_PSMGetLastErrorString()) EndIf
 ; ------------------ ; Handle login here!  ; ------------------ WinWait("vSphere Web Client - Windows Internet Explorer") BlockInput(1) Sleep("10000") Send($TargetUsername) Send("{TAB}") Send($TargetPassword) Send("{Enter}") BlockInput(0)
 ; Muthaheer: Terminate PSM Dispatcher utils wrapper LogWrite("Terminating Dispatcher Utils Wrapper") PSMGenericClient_Term()
 Return $PSM_ERROR_SUCCESSEndFunc
;==================================; Functions;==================================; #FUNCTION# ====================================================================================================================; Name...........: Error; Description ...: An exception handler - displays an error message and terminates the dispatcher; Parameters ....: $ErrorMessage - Error message to display;        $Code    - [Optional] Exit error code; ===============================================================================================================================Func Error($ErrorMessage, $Code = -1)
 ; If the dispatcher utils DLL was already initialized, write an error log message and terminate the wrapper if (PSMGenericClient_IsInitialized()) Then  LogWrite($ErrorMessage, True)  PSMGenericClient_Term() EndIf
 Local $MessageFlags = BitOr(0, 16, 262144) ; 0=OK button, 16=Stop-sign icon, 262144=MsgBox has top-most attribute set
 MsgBox($MessageFlags, $ERROR_MESSAGE_TITLE, $ErrorMessage)
 ; If the connection component was already invoked, terminate it if ($ConnectionClientPID <> 0) Then  ProcessClose($ConnectionClientPID)  $ConnectionClientPID = 0 EndIf
 Exit $CodeEndFunc
; #FUNCTION# ====================================================================================================================; Name...........: LogWrite; Description ...: Write a PSMWinSCPDispatcher log message to standard PSM log file; Parameters ....: $sMessage - [IN] The message to write;                  $LogLevel - [Optional] [IN] Defined if the message should be handled as an error message or as a trace messge; Return values .: $PSM_ERROR_SUCCESS - Success, otherwise error - Use PSMGenericClient_PSMGetLastErrorString for details.; ===============================================================================================================================Func LogWrite($sMessage, $LogLevel = $LOG_LEVEL_TRACE) Return PSMGenericClient_LogWrite($LOG_MESSAGE_PREFIX & $sMessage, $LogLevel)EndFunc
; #FUNCTION# ====================================================================================================================; Name...........: PSMGenericClient_GetSessionProperty; Description ...: Fetches properties required for the session; Parameters ....: None; Return values .: None; ===============================================================================================================================Func FetchSessionProperties() ; CHANGE_ME if (PSMGenericClient_GetSessionProperty("Username", $TargetUsername) <> $PSM_ERROR_SUCCESS) Then  Error(PSMGenericClient_PSMGetLastErrorString()) EndIf
 if (PSMGenericClient_GetSessionProperty("Password", $TargetPassword) <> $PSM_ERROR_SUCCESS) Then  Error(PSMGenericClient_PSMGetLastErrorString()) EndIf
EndFunc
