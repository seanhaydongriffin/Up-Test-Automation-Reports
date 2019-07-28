#include-once
#Include <Array.au3>
#include <GuiEdit.au3>
#include "cURL.au3"
#Region Header
#cs
	Title:   		Janison Insights Automation UDF Library for AutoIt3
	Filename:  		JanisonInsights.au3
	Description: 	A collection of functions for creating, attaching to, reading from and manipulating Janison Insights
	Author:   		seangriffin
	Version:  		V0.1
	Last Update: 	25/02/18
	Requirements: 	AutoIt3 3.2 or higher,
					Janison Insights Release x.xx,
					cURL xxx
	Changelog:		---------24/12/08---------- v0.1
					Initial release.
#ce
#EndRegion Header
#Region Global Variables and Constants
Global Const $sap_vkey[100] = [ "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", _
								"F11", _ ; NOTE - "F11" is the same as "CTRL+S"
								"F12", _ ; NOTE - "F12" is the same as "Esc"
								"Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5", "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9", _
								"Shift+Ctrl+0", "Shift+F11", "Shift+F12", _
								"Ctrl+F1", "Ctrl+F2", "Ctrl+F3", "Ctrl+F4", "Ctrl+F5", "Ctrl+F6", "Ctrl+F7", "Ctrl+F8", "Ctrl+F9", "Ctrl+F10", _
								"Ctrl+F11", "Ctrl+F12", _
								"Ctrl+Shift+F1", "Ctrl+Shift+F2", "Ctrl+Shift+F3", "Ctrl+Shift+F4", "Ctrl+Shift+F5", _
								"Ctrl+Shift+F6", "Ctrl+Shift+F7", "Ctrl+Shift+F8", "Ctrl+Shift+F9", "Ctrl+Shift+F10", "Ctrl+Shift+F11", _
								"Ctrl+Shift+F12", _
								"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
								"Ctrl+E", "Ctrl+F", "Ctrl+A", "Ctrl+D", "Ctrl+N", "Ctrl+O", "Shift+D", "Ctrl+I", "Shift+I", "Alt+B", _
								"Ctrl+Page up", "Page up", "Page down", "Ctrl+Page down", "Ctrl+G", "Ctrl+R", "Ctrl+P", _
								"", "", "", "", "", "", "", "Shift+F10", "", "", "", "", "" ]
Global $testrail_domain = ""
Global $testrail_username = ""
Global $testrail_password = ""
Global $testrail_json = ""
Global $testrail_html = ""
#EndRegion Global Variables and Constants
#Region Core functions
; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsSetup()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsSetup()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailSetup()

	; Initialise cURL
	cURL_initialise()


EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsShutdown()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsShutdown()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailShutdown()

	; Clean up cURL
	cURL_cleanup()

EndFunc


; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsDomainSet()
; Description ...:	Sets the domain to use in all other functions.
; Syntax.........:	_InsightsDomainSet($domain)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailDomainSet($domain)

	$testrail_domain = $domain
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsLogin()
; Description ...:	Login a user to Janison Insights.
; Syntax.........:	_InsightsLogin($username, $password)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailLogin($username, $password)

	$testrail_username = $username
	$testrail_password = $password
EndFunc

; Authentication


Func _TestRailAuth()

	Local $iPID = Run('curl.exe -k -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/auth/login -c cookies.txt -d "name=' & $testrail_username & '&password=' & $testrail_password & '&rememberme=1" -X POST', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit
EndFunc


; Projects

Func _TestRailGetProject($project_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_project/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

EndFunc

Func _TestRailGetProjectAnnouncement($project_id)

	_TestRailGetProject($project_id)
	$testrail_json = "{""project"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)
	Local $announcement = Json_Get($decoded_json, '.project.announcement')
	Return $announcement
EndFunc

Func _TestRailGetProjects()

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_projects", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	$testrail_json = $response[2]

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_projects', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc

Func _TestRailGetProjectsIDAndNameArray()

	Local $output[0][2]

	_TestRailGetProjects()

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"name":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $name = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$name = StringMid($name, StringInStr($name, ":", 0, -1) + 1)
		$name = StringReplace($name, """", "")
		Local $id_name = $id & "|" & $name
		_ArrayAdd($output, $id_name)
	Next

	Return $output
EndFunc

; Suites

Func _TestRailGetSuite($suite_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_suite/' & $suite_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

EndFunc

Func _TestRailGetSuitesIdName($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_suites/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_suites/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"name":"(.*)",', 3)
	Return $rr


EndFunc

Func _TestRailGetCases($project_id, $suite_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_cases/' & $project_id & '&suite_id=' & $suite_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
EndFunc

Func _TestRailGetCasesIdTitleReferences($project_id, $suite_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_cases/" & $project_id & "&suite_id=" & $suite_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_cases/' & $project_id & '&suite_id=' & $suite_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"title":"(.*)",.*"refs":"(.*)"', 3)
	Return $rr
EndFunc

Func _TestRailGetCase($case_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_case/" & $case_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailGetRunsIdName($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_runs/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_runs/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"name":"(.*)"', 3)
	Return $rr

EndFunc

Func _TestRailGetRun($run_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_run/" & $run_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailGetRunIDFromPlanIDAndRunName($plan_id, $run_name)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '"runs":\[{"id":\d+,.*"name":"' & $run_name & '"', 1)
	$rr[0] = StringLeft($rr[0], StringInStr($rr[0], ",") - 1)
	$rr[0] = StringMid($rr[0], StringInStr($rr[0], ":", 0, 2) + 1)

	Return $rr[0]


EndFunc

Func _TestRailGetPlanRunsID($plan_id)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),"suite_id":\d+,"name":".*","description"', 3)

	return $rr
EndFunc

Func _TestRailGetPlanRunsIDAndNameArray($plan_id)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),"suite_id":\d+,"name":"(.*)","description"', 3)

	return $rr
EndFunc


Func _TestRailGetResults($test_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_results/" & $test_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailGetResultsForCaseIdTestIdStatusIdDefects($run_id, $case_id)

;	Local $cmd = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	$rr = StringRegExp($testrail_json, '(?U)"defects":"(.*)","id":(.*),"status_id":(.*),"test_id":(.*),', 3)
	Return $rr


EndFunc

Func _TestRailGetResultsIdStatusIdDefects($test_id)

;	Local $cmd = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results/' & $test_id & '&limit=1', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	filedelete("D:\dwn\fred.txt")
	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"test_id":(.*),.*"status_id":(.*),.*"defects":(.*),"custom_step_results"', 3)
;	_ArrayDisplay($rr)
;	Exit
	Return $rr


EndFunc

Func _TestRailGetResultsForRunIdStatusIdCreatedOnDefects($run_id)

;	Local $cmd = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id & '&limit=1', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	filedelete("D:\dwn\fred.txt")
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"test_id":(.*),.*"status_id":(.*),.*"created_on":(.*),.*"defects":(.*),"custom_step_results"', 3)
;	_ArrayDisplay($rr)
;	Exit
	Return $rr


EndFunc

Func _TestRailGetResultsForRun($run_id, $created_after, $created_before)

	if StringLen($created_after) > 0 Then

		$created_after = "&created_after=" & $created_after
	EndIf

	if StringLen($created_before) > 0 Then

		$created_before = "&created_before=" & $created_before
	EndIf

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id & $created_after & $created_before, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc

Func _TestRailGetTests($run_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_tests/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

EndFunc

Func _TestRailGetTestsIdTitleCaseIdRunId($run_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_tests/" & $run_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_tests/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"case_id":(.*),.*"run_id":(.*),.*"title":"(.*)",', 3)
	Return $rr


EndFunc

Func _TestRailGetTestsTitleAndIDFromRunID($run_id)

	Local $test_title_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetTests($run_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"title":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $title = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$title = StringMid($title, StringInStr($title, ":", 0, -1) + 1)
		$title = StringReplace($title, """", "")
		$test_title_and_id_dict.Add($title, $id)
	Next

	Return $test_title_and_id_dict

EndFunc

Func _TestRailGetTestsReferenceAndIDFromRunID($run_id)

	Local $test_refs_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetTests($run_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"refs":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $refs = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $id = ' & $id & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		$refs = StringMid($refs, StringInStr($refs, ":", 0, -1) + 1)
		$refs = StringReplace($refs, """", "")
;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $refs = ' & $refs & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		$test_refs_and_id_dict.Add($refs, $id)
	Next

	Return $test_refs_and_id_dict

EndFunc

Func _TestRailGetPlan($plan_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_plan/" & $plan_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_plan/' & $plan_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc

Func _TestRailGetPlans($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_plans/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_plans/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


EndFunc

Func _TestRailGetResultsElapsedForRuns($run_ids, $created_after, $created_before)

	Local $output[0][2]
	Local $run_id = StringSplit($run_ids, ",", 3)

	for $i = 0 to (UBound($run_id) - 1)

		_TestRailGetResultsForRun($run_id[$i], $created_after, $created_before)

		$testrail_json = "{""results"":" & $testrail_json & "}"
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		Local $decoded_json = Json_Decode($testrail_json)

		for $j = 0 to 99999

			Local $elapsed = Json_Get($decoded_json, '.results[' & $j & '].elapsed')

			if @error > 0 or StringLen($elapsed) < 1 Then ExitLoop

			Local $created_on = Json_Get($decoded_json, '.results[' & $j & '].created_on')
			Local $created_on_date = _EPOCH_decrypt($created_on)

			_ArrayAdd($output, $created_on_date & Chr(28) & $elapsed, 0, Chr(28), @CRLF, 1)
		Next
	Next

;	_ArrayDisplay($output)
	Return $output
EndFunc

Func _TestRailGetCreatedOnTestSavedRuntimeForRuns($project_id, $run_suite, $created_after, $created_before, $testcase_complexity_execution_weightage_dict)

	Local $output[0][3]

	;_ArrayDisplay($run_suite)

	for $i = 0 to (UBound($run_suite) - 1)

		_TestRailGetResultsForRun($run_suite[$i][0], $created_after, $created_before)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $run_suite[$i][0] = ' & $run_suite[$i][0] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $created_after = ' & $created_after & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $created_before = ' & $created_before & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		$testrail_json = "{""results"":" & $testrail_json & "}"
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		Local $results_decoded_json = Json_Decode($testrail_json)

		if StringLen($testrail_json) > StringLen("{""results"":[]}") Then

			_TestRailGetTests($run_suite[$i][0])
			$testrail_json = "{""tests"":" & $testrail_json & "}"
			Local $tests_decoded_json = Json_Decode($testrail_json)

			_TestRailGetCases($project_id, $run_suite[$i][1])
			$testrail_json = "{""cases"":" & $testrail_json & "}"
			Local $cases_decoded_json = Json_Decode($testrail_json)
		EndIf


		for $j = 0 to 99999

			Local $elapsed = Json_Get($results_decoded_json, '.results[' & $j & '].elapsed')

			if @error > 0 Then ExitLoop

			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $elapsed = ' & $elapsed & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			if StringLen($elapsed) > 0 Then

				; convert testrail elapsed into seconds elapsed

				Local $elapsed_part = StringSplit($elapsed, " ", 3)
				Local $elapsed_seconds = 0

				for $each in $elapsed_part

					if StringInStr($each, "h") > 0 Then

						$each = StringReplace($each, "h", "")
						$elapsed_seconds = $elapsed_seconds + ($each * 3600)
					EndIf

					if StringInStr($each, "m") > 0 Then

						$each = StringReplace($each, "m", "")
						$elapsed_seconds = $elapsed_seconds + ($each * 60)
					EndIf

					if StringInStr($each, "s") > 0 Then

						$each = StringReplace($each, "s", "")
						$elapsed_seconds = $elapsed_seconds + $each
					EndIf
				Next

				Local $created_on = Json_Get($results_decoded_json, '.results[' & $j & '].created_on')
				Local $created_on_date = _EPOCH_decrypt($created_on)
				Local $test = Json_Get($results_decoded_json, '.results[' & $j & '].test_id')
				Local $saved_runtime = 0

				; find the title and complexity of the test

				for $k = 0 to 99999

					Local $test_id = Json_Get($tests_decoded_json, '.tests[' & $k & '].id')

					if @error > 0 Then ExitLoop

					if Number($test_id) = Number($test) Then

						$test = Json_Get($tests_decoded_json, '.tests[' & $k & '].title')
						$test = StringReplace($test, "'", "''")
						Local $test_case_id = Json_Get($tests_decoded_json, '.tests[' & $k & '].case_id')

						for $l = 0 to 99999

							Local $case_id = Json_Get($cases_decoded_json, '.cases[' & $l & '].id')

							if @error > 0 Then ExitLoop

							if Number($case_id) = Number($test_case_id) Then

								$complexity = Json_Get($cases_decoded_json, '.cases[' & $l & '].custom_complexity')
								Local $manual_vs_automated_multiplier = $testcase_complexity_execution_weightage_dict.Item($complexity)
								$saved_runtime = int((($elapsed_seconds * Number($manual_vs_automated_multiplier)) - $elapsed_seconds) / 60)
								ExitLoop 2
							EndIf
						Next
					EndIf
				Next

				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $test = ' & $test & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $complexity = ' & $complexity & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

				_ArrayAdd($output, $created_on_date & Chr(28) & $test & Chr(28) & $saved_runtime, 0, Chr(28), @CRLF, 1)
				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $created_on_date & Chr(28) & $test_id & Chr(28) & $elapsed_seconds = ' & $created_on_date & Chr(28) & $test & Chr(28) & $saved_runtime & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			EndIf
		Next
	Next

;	_ArrayDisplay($output)
	Return $output
EndFunc



Func _TestRailGetPlansIDAndNameArray($project_id)

	Local $output[0][2]

	_TestRailGetPlans($project_id)

	$testrail_json = "{""plans"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.plans[' & $i & '].id')

		if @error > 0 Then ExitLoop

		Local $name = Json_Get($decoded_json, '.plans[' & $i & '].name')

		_ArrayAdd($output, $id & Chr(28) & $name, 0, Chr(28), @CRLF, 1)
	Next

;	_ArrayDisplay($output)
	Return $output
EndFunc

Func _TestRailGetPlansRunsIDArray($plan_ids)

	Local $output[0]
	Local $plan_id = StringSplit($plan_ids, ",", 3)

	for $i = 0 to (UBound($plan_id) - 1)

		_TestRailGetPlan($plan_id[$i])

		$testrail_json = "{""plan"":" & $testrail_json & "}"
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		Local $decoded_json = Json_Decode($testrail_json)

		for $j = 0 to 99999

			Local $entry_id = Json_Get($decoded_json, '.plan.entries[' & $j & '].id')

			if @error > 0 Then ExitLoop

			for $k = 0 to 99999

				Local $id = Json_Get($decoded_json, '.plan.entries[' & $j & '].runs[' & $k & '].id')
				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $id = ' & $id & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

				if @error > 0 Or StringLen($id) < 1 Then ExitLoop

				_ArrayAdd($output, $id, 0, Chr(28), @CRLF, 1)
			Next
		Next
	Next

;	_ArrayDisplay($output)
	Return $output
EndFunc

Func _TestRailGetPlansRunsIDSuitesIDArray($plan_ids)

	Local $output[0][2]
	Local $plan_id = StringSplit($plan_ids, ",", 3)

	for $i = 0 to (UBound($plan_id) - 1)

		_TestRailGetPlan($plan_id[$i])

		$testrail_json = "{""plan"":" & $testrail_json & "}"
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		Local $decoded_json = Json_Decode($testrail_json)

		for $j = 0 to 99999

			Local $entry_id = Json_Get($decoded_json, '.plan.entries[' & $j & '].id')

			if @error > 0 Then ExitLoop

			for $k = 0 to 99999

				Local $id = Json_Get($decoded_json, '.plan.entries[' & $j & '].runs[' & $k & '].id')
				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $id = ' & $id & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

				if @error > 0 Or StringLen($id) < 1 Then ExitLoop

				Local $suite_id = Json_Get($decoded_json, '.plan.entries[' & $j & '].runs[' & $k & '].suite_id')

				_ArrayAdd($output, $id & Chr(28) & $suite_id, 0, Chr(28), @CRLF, 1)
			Next
		Next
	Next

;	_ArrayDisplay($output)
	Return $output
EndFunc

Func _TestRailGetPlanIDByName($project_id, $plan_name)

	_TestRailGetPlans($project_id)

	$rr = StringRegExp($testrail_json, '"id":\d+,"name":"' & $plan_name & '"', 1)
	$rr[0] = StringLeft($rr[0], StringInStr($rr[0], ",") - 1)
	$rr[0] = StringMid($rr[0], StringInStr($rr[0], ":") + 1)

	Return $rr[0]
EndFunc

Func _TestRailAddResult($test_id, $status_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/add_result/" & $test_id, "", 0, 0, "", "Content-Type: application/json", '{"status_id":' & $status_id & '}', 0, 1, 0, $testrail_username & ":" & $testrail_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _TestRailAddResults($run_id, $results_arr)

	; create the results JSON to post from the results array

	Local $results_json = '{"results":['

	for $i = 0 to (UBound($results_arr) - 1) step 3

		if StringLen($results_json) > StringLen('{"results":[') Then

			$results_json = $results_json & ','
		EndIf

		$results_json = $results_json & '{"test_id":' & $results_arr[$i + 0] & ',"status_id":' & $results_arr[$i + 1] & ',"comment":"' & $results_arr[$i + 2] & '"}'
	Next

	$results_json = $results_json & ']}'
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $results_json = ' & $results_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	FileDelete(@ScriptDir & "\curl_in.json")
	FileWrite(@ScriptDir & "\curl_in.json", $results_json)
	Local $iPID = Run('curl.exe -s -k -H "Content-Type: application/json" --data @curl_in.json -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/add_results/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	; unfortunately the below is unreliable.  Working intermittently
;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/add_results/" & $run_id, "", 0, 0, "", "Content-Type: application/json", $results_json, 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $response[2] = ' & $response[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc


Func _TestRailGetIdFromTitle($json, $title)

	$rr = StringRegExp($json, '"id":.*"title":"' & $title & '"', 1)
	$tt = StringMid($rr[0], StringLen('"id":') + 1, StringInStr($rr[0], ",") - (StringLen('"id":') + 1))
	Return $tt

EndFunc

Func _TestRailGetStatusesIdLabel()

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_statuses", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_statuses/', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	filedelete("D:\dwn\fred.txt")
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"label":"(.*)"', 3)
	Return $rr


EndFunc

Func _TestRailGetStatusLabelAndID()

	Local $status_label_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetStatuses()

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"label":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $label = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$label = StringMid($label, StringInStr($label, ":", 0, -1) + 1)
		$label = StringReplace($label, """", "")

		$status_label_and_id_dict.Add($label, $id)
	Next

	Return $status_label_and_id_dict

;	_ArrayDisplay($rr)

EndFunc


Func _TestRailGetSections($project_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_sections/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	$testrail_json = $response[2]
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc


Func _TestRailGetSectionNameAndDepth($project_id)

	_TestRailGetSections($project_id)

	$rr = StringRegExp($testrail_json, '(?U)"name":"(.*)",.*"depth":(\d+)', 3)

	Return $rr

EndFunc

; Jira integration



Func _TestRailGetTestCases($key)

	$response = cURL_easy($testrail_domain & "/index.php?/ext/jira/render_panel&ae=connect&av=1&issue=" & $key & "&panel=references&login=button&frame=tr-frame-panel-references", "cookies.txt", 1, 0, "", "Content-Type: text/html; charset=UTF-8", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	$testrail_html = $response[2]
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_html = ' & $testrail_html & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc


Func _EPOCH_decrypt($iEpochTime)
    Local $iDayToAdd = Int($iEpochTime / 86400)
    Local $iTimeVal = Mod($iEpochTime, 86400)
    If $iTimeVal < 0 Then
        $iDayToAdd -= 1
        $iTimeVal += 86400
    EndIf
    Local $i_wFactor = Int((573371.75 + $iDayToAdd) / 36524.25)
    Local $i_xFactor = Int($i_wFactor / 4)
    Local $i_bFactor = 2442113 + $iDayToAdd + $i_wFactor - $i_xFactor
    Local $i_cFactor = Int(($i_bFactor - 122.1) / 365.25)
    Local $i_dFactor = Int(365.25 * $i_cFactor)
    Local $i_eFactor = Int(($i_bFactor - $i_dFactor) / 30.6001)
    Local $aDatePart[3]
    $aDatePart[2] = $i_bFactor - $i_dFactor - Int(30.6001 * $i_eFactor)
    $aDatePart[1] = $i_eFactor - 1 - 12 * ($i_eFactor - 2 > 11)
    $aDatePart[0] = $i_cFactor - 4716 + ($aDatePart[1] < 3)
    Local $aTimePart[3]
    $aTimePart[0] = Int($iTimeVal / 3600)
    $iTimeVal = Mod($iTimeVal, 3600)
    $aTimePart[1] = Int($iTimeVal / 60)
    $aTimePart[2] = Mod($iTimeVal, 60)

	$ST = _Date_Time_EncodeSystemTime($aDatePart[1], $aDatePart[2], $aDatePart[0], $aTimePart[0], $aTimePart[1], $aTimePart[2])
	$LT = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($ST))
;	ConsoleWrite(_Date_Time_SystemTimeToDateTimeStr($LT,1) & @CRLF)
	Return _Date_Time_SystemTimeToDateTimeStr($LT,1)

 ;   Return StringFormat("%.2d/%.2d/%.2d %.2d:%.2d:%.2d", $aDatePart[0], $aDatePart[1], $aDatePart[2], $aTimePart[0], $aTimePart[1], $aTimePart[2])
EndFunc

