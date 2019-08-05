#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <Date.au3>
#include <File.au3>
#include <Array.au3>
#Include "Json.au3"
#include "Jira.au3"
#include "Confluence.au3"
#include "TestRail.au3"
#include <WindowsConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Crypt.au3>
#include "Toast.au3"

Global $app_name = "UP Test Automation Reports"
Global $jira_project_name = "INS"
Global $confluence_page_key = "469172334"
Global $confluence_ancestor_key = "469172330"
Global $testrail_project_id = 49
Global $github_release_url = "https://github.com/seanhaydongriffin/UP-Test-Automation-Reports/releases/download/v0.1/UP.Test.Automation.Reports.portable.exe"

Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"
Global $log_filepath = @ScriptDir & "\" & $app_name & ".log"
Global $html, $markup, $storage_format
Global $aResult, $iRows, $iColumns, $iRval, $run_name = "", $max_num_defects = 0, $max_num_days = 0, $version_name = ""

; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)

FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE SavedRuntime (CreatedOn,Test,SavedRuntime integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SavedCases (Test,SavedCases integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Bug (Key,Created,Priority);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SubTask (Key,Summary,Description,StoryKey,FixVersion,Status,EstimatedTime,TimeSpent,ProgressPercent,Environment);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SubTaskStateHistory (Key,Date,Status);") ; CREATE a Table


; TestRail

_TestRailDomainSet("https://janison.testrail.com")

_Toast_Set(0, -1, -1, -1, -1, -1, "", 100, 100)
_Toast_Show(0, $app_name, "Login to TestRail ...", -30, False, True)

Local $testrail_username = IniRead($ini_filename, "main", "username", "")
Local $testrail_encrypted_password = IniRead($ini_filename, "main", "password", "")
Global $testrail_decrypted_password = ""

if stringlen($testrail_encrypted_password) > 0 Then

	$testrail_decrypted_password = _Crypt_DecryptData($testrail_encrypted_password, @ComputerName & @UserName, $CALG_AES_256)
	$testrail_decrypted_password = BinaryToString($testrail_decrypted_password)
EndIf

if stringlen($testrail_decrypted_password) > 0 Then

	_TestRailLogin($testrail_username, $testrail_decrypted_password)
	_TestRailAuth()
EndIf

if stringlen($testrail_decrypted_password) = 0 or StringInStr($testrail_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

	_Toast_Show(0, $app_name, "Username or password incorrect or not set.                       " & @CRLF & "Set your TestRail login below." & @CRLF & @CRLF & @CRLF & @CRLF & @CRLF, -9999, False, True)
	GUICtrlCreateLabel("Username:", 10, 70, 80, 20)
	Local $username_input = GUICtrlCreateInput("", 80, 70, 200, 20)
	GUICtrlCreateLabel("Password:", 10, 90, 80, 20)
	Local $password_input = GUICtrlCreateInput("", 80, 90, 200, 20, $ES_PASSWORD)
	$done_button = GUICtrlCreateButton("Done", 80, 110, 80, 20)

	While 1

		$msg = GUIGetMsg()

		if $msg = $done_button Then

			Local $tmp_username = GUICtrlRead($username_input)
			Local $tmp_password = GUICtrlRead($password_input)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $tmp_password = ' & $tmp_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			_Toast_Show(0, $app_name, "Login to TestRail ...", -30, False, True)
			_TestRailLogin($tmp_username, $tmp_password)
			_TestRailAuth()

			if StringInStr($testrail_json, "Email/Login or Password is incorrect", 1) = 0 Then

				IniWrite($ini_filename, "main", "username", $tmp_username)
				$testrail_encrypted_password = _Crypt_EncryptData($tmp_password, @ComputerName & @UserName, $CALG_AES_256)
				IniWrite($ini_filename, "main", "password", $testrail_encrypted_password)
			EndIf

			_Toast_Hide()
			ExitLoop
		EndIf

		if $hToast_Handle = 0 Then

			Exit
		EndIf
	WEnd
EndIf

if StringInStr($testrail_json, "Email/Login or Password is incorrect", 1) > 0 Then

	_Toast_Show(0, $app_name, "Username or password incorrect or not set." & @CRLF & "Exiting ...", -5, true, True)
	Exit
EndIf

; Confluence

_JiraSetup()
_JiraDomainSet("https://janisoncls.atlassian.net")
_Toast_Set(0, -1, -1, -1, -1, -1, "", 100, 100)
_Toast_Show(0, $app_name, "Login to Confluence ...", -30, False, True)

Local $confluence_username = IniRead($ini_filename, "main", "confluenceusername", "")
Local $confluence_encrypted_password = IniRead($ini_filename, "main", "confluencepassword", "")
Global $confluence_decrypted_password = ""

if stringlen($confluence_encrypted_password) > 0 Then

	$confluence_decrypted_password = _Crypt_DecryptData($confluence_encrypted_password, @ComputerName & @UserName, $CALG_AES_256)
	$confluence_decrypted_password = BinaryToString($confluence_decrypted_password)
EndIf

if stringlen($confluence_decrypted_password) > 0 Then

	_JiraLogin($confluence_username, $confluence_decrypted_password)
	_JiraGetCurrentUser()
EndIf

if stringlen($confluence_decrypted_password) = 0 or StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

	_Toast_Show(0, $app_name, "Username or password incorrect or not set.                       " & @CRLF & "Set your Confluence login below." & @CRLF & @CRLF & @CRLF & @CRLF & @CRLF, -9999, False, True)
	GUICtrlCreateLabel("Username:", 10, 70, 80, 20)
	Local $username_input = GUICtrlCreateInput("", 80, 70, 200, 20)
	GUICtrlCreateLabel("Password:", 10, 90, 80, 20)
	Local $password_input = GUICtrlCreateInput("", 80, 90, 200, 20, $ES_PASSWORD)
	$done_button = GUICtrlCreateButton("Done", 80, 110, 80, 20)

	While 1

		$msg = GUIGetMsg()

		if $msg = $done_button Then

			$confluence_username = GUICtrlRead($username_input)
			Local $confluence_decrypted_password = GUICtrlRead($password_input)
			_Toast_Show(0, $app_name, "Login to Confluence ...", -30, False, True)
			_JiraLogin($confluence_username, $confluence_decrypted_password)
			_JiraGetCurrentUser()

			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			if StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) = 0 Then

				IniWrite($ini_filename, "main", "confluenceusername", $confluence_username)
				$confluence_encrypted_password = _Crypt_EncryptData($confluence_decrypted_password, @ComputerName & @UserName, $CALG_AES_256)
				IniWrite($ini_filename, "main", "confluencepassword", $confluence_encrypted_password)
			EndIf

			_Toast_Hide()
			ExitLoop
		EndIf

		if $hToast_Handle = 0 Then

			Exit
		EndIf
	WEnd
EndIf

if StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

	_Toast_Show(0, $app_name, "Username or password incorrect or not set." & @CRLF & "Exiting ...", -5, true, True)
	Exit
EndIf






; get the announcement of the project (which contains the testcase complexity weightages)

Local $project_announcement = _TestRailGetProjectAnnouncement($testrail_project_id)
ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $project_announcement = ' & $project_announcement & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

Local $project_announcement_line = StringSplit($project_announcement, @crlf, 3)
Local $testcase_complexity_execution_weightage_found = False
Local $testcase_complexity_execution_weightage_index = 0
Global $testcase_complexity_execution_weightage_dict = ObjCreate("Scripting.Dictionary")

for $each in $project_announcement_line

	if $testcase_complexity_execution_weightage_found = True Then

		if StringInStr($each, "||") = 0 Then

			$testcase_complexity_execution_weightage_found = False
		Else

			$testcase_complexity_execution_weightage_index = $testcase_complexity_execution_weightage_index + 1
			Local $table_row_cell = StringSplit($each, "|", 3)
			$table_row_cell[2] = StringStripWS($table_row_cell[2], 3)
			$table_row_cell[3] = StringStripWS($table_row_cell[3], 3)
			$table_row_cell[4] = StringStripWS($table_row_cell[4], 3)
			$testcase_complexity_execution_weightage_dict.Add ($testcase_complexity_execution_weightage_index, $table_row_cell[4])
		EndIf
	EndIf

	if StringInStr($each, "||| :Complexity | :Automated Runtime Approx (mins) | :Manual Vs Automated Multiplier") > 0 Then

		$testcase_complexity_execution_weightage_found = True
	EndIf
Next


; get all plans for the project

_Toast_Show(0, $app_name, "get all plans for the project", -30, False, True)
Local $plan = _TestRailGetPlansIDAndNameArray($testrail_project_id)

; for every SIT plan get the runs

Local $plan_ids = ""

for $i = 0 to (UBound($plan) - 1)

	if StringInStr($plan[$i][1], "UP ") = 1 and StringInStr($plan[$i][1], " SIT") > 0 Then

		if StringLen($plan_ids) > 0 Then

			$plan_ids = $plan_ids & ","
		EndIf

		$plan_ids = $plan_ids & $plan[$i][0]
	EndIf
Next

ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $plan_ids = ' & $plan_ids & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


; for every plan get the runs

_Toast_Show(0, $app_name, "get all runs for the plans", -30, False, True)
Local $run_suite = _TestRailGetPlansRunsIDSuitesIDArray($plan_ids)

; for every run get the elapsed results

_Toast_Show(0, $app_name, "get all runtime saved from the runs", -30, False, True)

Local $todays_day_of_week = _DateToDayOfWeek(@YEAR, @MON, @MDAY)
Local $number_of_days_to_last_monday = 2 - $todays_day_of_week - 7
;Local $previous_quarter_date = StringLeft(_DateAdd("D", $number_of_days_to_last_monday - (7 * 11), _NowCalc()), 10)
;Local $previous_six_months_date = StringLeft(_DateAdd("D", $number_of_days_to_last_monday - (7 * 25), _NowCalc()), 10)
;$created_after_date = _DateDiff( 's',"1970/01/01 00:00:00", $previous_quarter_date & " 00:00:00")
;$created_after_date = _DateDiff( 's',"1970/01/01 00:00:00", $previous_six_months_date & " 00:00:00")


; RUN TIME SAVED REPORT


Local $tomorrow_date = StringLeft(_DateAdd("D", 1, _NowCalc()), 10)
$created_before_date = _DateDiff( 's',"1970/01/01 00:00:00", $tomorrow_date & " 00:00:00")
;Local $saved_runtime = _TestRailGetCreatedOnTestSavedRuntimeForRuns($testrail_project_id, $run_suite, $created_after_date, $created_before_date, $testcase_complexity_execution_weightage_dict)
Local $saved_runtime = _TestRailGetCreatedOnTestSavedRuntimeForRuns($testrail_project_id, $run_suite, "", $created_before_date, $testcase_complexity_execution_weightage_dict)
;Local $saved_runtime = _TestRailGetCreatedOnTestSavedRuntimeForRuns($testrail_project_id, $run_suite, "", "", $testcase_complexity_execution_weightage_dict)

for $i = 0 to (UBound($saved_runtime) - 1)

	$saved_runtime[$i][0] = StringLeft($saved_runtime[$i][0], 10)

	$query = "INSERT INTO SavedRuntime (CreatedOn,Test,SavedRuntime) VALUES ('" & $saved_runtime[$i][0] & "','" & $saved_runtime[$i][1] & "'," & $saved_runtime[$i][2] & ");"
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

_Toast_Show(0, $app_name, "Creating Run Time Saved Report", -30, False, True)

$storage_format = $storage_format & '<a href=\"' & $github_release_url & '\">Click to update this page</a><br /><br />'

;Local $query = "select strftime('%W', CreatedOn) as `Week Starting`, sum(SavedRuntime) as `Minutes Saved` from SavedRuntime group by `Week Starting`;"
;SQLite_with_weeks_to_Confluence_Chart("bar", "Run Time Saved (weekly) with Automation", "", "Week Starting", "Minutes Saved", "horizontal", "true", "false", "800", "480", "false", "", "after", "1,2", "1,2", "vertical", "", "", "day", "", "", "true", "", "", "", "", "", $query)
Local $query = "select `Week Starting`, (select sum(minutessaved) from (select strftime('%W', CreatedOn) as `Week Starting`, sum(SavedRuntime) as minutessaved from SavedRuntime group by `Week Starting`) c2 where c2.`Week Starting` <= c1.`Week Starting`) / 60 as `Hours Saved To Date` from (select strftime('%W', CreatedOn) as `Week Starting`, sum(SavedRuntime) as minutessaved from SavedRuntime group by `Week Starting`) c1;"
SQLite_with_weeks_to_Confluence_Chart("bar", "Manual Run Time Saved (weekly to date) with Automation", "Note - data prior to 2019/07/08 is not reportable", "Week Starting", "Hours Saved To Date", "horizontal", "false", "false", "800", "480", "false", "", "after", "1,2", "1,2", "vertical", "", "", "day", "", "", "true", "", "", "", "", "", $query)


; MANUAL TEST CASES SAVED REPORT

_Toast_Show(0, $app_name, "Get all manual test cases saved", -30, False, True)

Local $suite_id = _TestRailGetSuitesId($testrail_project_id, "UP - ")

for $i = 0 to (UBound($suite_id) - 1)

	Local $test_cases_saved = _TestRailGetCasesAutomationTestCasesSaved($testrail_project_id, $suite_id[$i])

	for $j = 0 to (UBound($test_cases_saved) - 1)

		if StringLen($test_cases_saved[$j][1]) > 0 Then

			$query = "INSERT INTO SavedCases (Test,SavedCases) VALUES ('" & $test_cases_saved[$j][0] & "'," & $test_cases_saved[$j][1] & ");"
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			_SQLite_Exec(-1, $query) ; INSERT Data
		EndIf
	Next
Next

_Toast_Show(0, $app_name, "Creating Manual Test Cases Saved Report", -30, False, True)

_SQLite_GetTable2d(-1, "select sum(SavedCases) from SavedCases;", $aResult, $iRows, $iColumns)
Local $query = "select Test as `Test Case`, SavedCases as `Test Cases Saved` from SavedCases;"
SQLite_to_Confluence_Chart("pie", "Manual Test Cases Saved (to date) with Automation", "Total = " & $aResult[1][0], "", "Test Cases Saved", "false", "false", "800", "480", "false", "", "", "1,2", "1,2", "vertical", "", "", "day", "", "", "true", "", "", "", "", $query)


; AUTOMATION BUGS REPORT

_Toast_Show(0, $app_name, "Get all bugs raised by automation", -30, False, True)

;$bug = _JiraGetSearchResultKeysCreatedDatePriorities("project = " & $jira_project_name & " AND issuetype = Bug AND labels = Automation AND labels = UP AND created >= " & StringReplace($previous_six_months_date, "/", "-"))
$bug = _JiraGetSearchResultKeysCreatedDatePriorities("project = " & $jira_project_name & " AND issuetype = Bug AND labels = Automation")

for $i = 0 to (UBound($bug) - 1) Step 3

	$query = "INSERT INTO Bug (Key,Created,Priority) VALUES ('" & $bug[$i] & "','" & $bug[$i+1] & "','" & $bug[$i+2] & "');"
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

_Toast_Show(0, $app_name, "Creating Automation Bugs Report", -30, False, True)

;Local $query = "select aaa as Week, (select count(Key) as `Blocker` from Bug where strftime('%W', Created) = aaa and Priority = 'Blocker') as `Blocker`, (select count(Key) as `Critical` from Bug where strftime('%W', Created) = aaa and Priority = 'Critical') as `Critical`, (select count(Key) as `Major` from Bug where strftime('%W', Created) = aaa and Priority = 'Major') as `Major`, (select count(Key) as `Minor` from Bug where strftime('%W', Created) = aaa and Priority = 'Minor') as `Minor` from (select strftime('%W', Created) as aaa from Bug) group by aaa;"
;SQLite_with_weeks_to_Confluence_Chart("bar", "Bugs Found (weekly) with Automation", "", "Week Starting", "Bugs Found", "horizontal", "true", "true", "800", "480", "false", "", "", "1,2,3,4,5", "1,2,3,4,5", "vertical", "", "", "day", "", "", "true", "", "", "", "1", "", $query)
Local $query = "select week, (select sum(blocker) from (select week, (select count(Key) as blocker from Bug where strftime('%W', Created) = week and Priority = 'Blocker') as Blocker from (select strftime('%W', Created) as week from Bug) group by week) b2 where b2.week <= b1.week) as Blocker, (select sum(critical) from (select week, (select count(Key) as critical from Bug where strftime('%W', Created) = week and Priority = 'Critical') as Critical from (select strftime('%W', Created) as week from Bug) group by week) c2 where c2.week <= b1.week) as Critical, (select sum(major) from (select week, (select count(Key) as major from Bug where strftime('%W', Created) = week and Priority = 'Major') as Major from (select strftime('%W', Created) as week from Bug) group by week) ma where ma.week <= b1.week) as Major, (select sum(minor) from (select week, (select count(Key) as minor from Bug where strftime('%W', Created) = week and Priority = 'Minor') as Minor from (select strftime('%W', Created) as week from Bug) group by week) mi where mi.week <= b1.week) as Minor from (select week, (select count(Key) as blocker from Bug where strftime('%W', Created) = week and Priority = 'Blocker') as blocker, (select count(Key) as critical from Bug where strftime('%W', Created) = week and Priority = 'Critical') as critical from (select strftime('%W', Created) as week from Bug) group by week) b1;"
SQLite_with_weeks_to_Confluence_Chart("bar", "Defects Raised (weekly to date) with Automation", "", "Week Starting", "Bugs Found To Date", "horizontal", "true", "true", "800", "480", "false", "", "after", "1,2,3,4,5", "1,2,3,4,5", "vertical", "", "", "day", "", "", "true", "", "", "", "1", "", $query)


; CORE COVERAGE REPORT


_Toast_Show(0, $app_name, "get all sub-tasks for the project", -30, False, True)
$issue = _JiraGetSearchResultKeysSummariesIssueTypeNameStoryKeyRequirements2("summary,description,issuetype,parent,labels,fixVersions,status,aggregateprogress,environment", "project = QA AND issuetype = Sub-task AND labels in (Core) AND labels in (Automation) AND labels in (Assessments)")
;_ArrayDisplay($issue)
;Exit

for $i = 0 to (UBound($issue) - 1) Step 13

	$issue[$i + 1] = StringReplace($issue[$i + 1], "'", "''")
	$issue[$i + 2] = StringReplace($issue[$i + 2], "'", "''")

	$query = "INSERT INTO SubTask (Key,Summary,Description,StoryKey,FixVersion,Status,EstimatedTime,TimeSpent,ProgressPercent,Environment) VALUES ('" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 2] & "','" & $issue[$i + 4] & "','" & $issue[$i + 6] & "','" & $issue[$i + 7] & "','" & $issue[$i + 8] & "','" & $issue[$i + 9] & "','" & $issue[$i + 10] & "','" & $issue[$i + 11] & "');"
	_FileWriteLog($log_filepath, "SubTask " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
	_SQLite_Exec(-1, $query) ; INSERT Data

	if StringLen($issue[$i + 12]) > 0 Then

		Local $status_history = StringSplit($issue[$i + 12], "~", 3)

		for $j = 0 to (UBound($status_history) - 1)

			Local $status_history_part = StringSplit($status_history[$j], ",", 3)

			$query = "INSERT INTO SubTaskStateHistory (Key,Date,Status) VALUES ('" & $issue[$i] & "','" & $status_history_part[0] & "','" & $status_history_part[1] & "');"
			_FileWriteLog($log_filepath, "SubTask " & ($i + 1) & " of " & UBound($issue) & " = " & $query)
			_SQLite_Exec(-1, $query) ; INSERT Data
		Next
	EndIf
Next

Local $num_coverage = 0
Local $num_coverage_skipped = 0
Local $num_coverage_yes = 0
Local $num_coverage_no = 0
Local $pcnt_coverage_yes = 0

$iRval = _SQLite_GetTable2d(-1, "select count(*) from SubTask;", $aResult, $iRows, $iColumns)

If $iRval = $SQLITE_OK Then

	$num_coverage = $aResult[1][0]
EndIf

$iRval = _SQLite_GetTable2d(-1, "select count(*) from SubTask where Status = 'Done';", $aResult, $iRows, $iColumns)

If $iRval = $SQLITE_OK Then

	$num_coverage_skipped = $aResult[1][0]
EndIf

$num_coverage = $num_coverage - $num_coverage_skipped
ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $num_coverage = ' & $num_coverage & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;Local $query = "select week, printf(""%.1f"", ((count(week) * 1.0) / " & $num_coverage & " * 100)) as pcnt from (select Key, strftime('%W', Date) as week, Status from SubTaskStateHistory group by Key, strftime('%W', Date) having SubTaskStateHistory.rowid = min(SubTaskStateHistory.rowid)) where Status in ('Waiting For Build', 'In Progress', 'Beta Testing', 'Ready', 'Closed', 'Cancelled', 'Resolved', 'Test Run') group by week;"
;SQLite_with_weeks_to_Confluence_Chart("bar", "Core Coverage (weekly) with Automation", "", "Week Starting", "Percent of Functions Covered", "horizontal", "false", "true", "800", "480", "false", "", "after", "1,2", "1,2", "vertical", "", "", "day", "", "", "true", "", "", "", "1", "", $query)
Local $query = "select week, printf(""%.1f"", ((select sum(number) from (select week, count(week) as number from (select Key, strftime('%W', Date) as week, Status from SubTaskStateHistory group by Key, strftime('%W', Date) having SubTaskStateHistory.rowid = min(SubTaskStateHistory.rowid)) where Status in ('Waiting For Build', 'In Progress', 'Beta Testing', 'Ready', 'Closed', 'Cancelled', 'Resolved', 'Test Run') group by week) c2 where c2.week <= c1.week) * 1.0 / " & $num_coverage & " * 100)) as pcnt from (select week, count(week) as number from (select Key, strftime('%W', Date) as week, Status from SubTaskStateHistory group by Key, strftime('%W', Date) having SubTaskStateHistory.rowid = min(SubTaskStateHistory.rowid)) where Status in ('Waiting For Build', 'In Progress', 'Beta Testing', 'Ready', 'Closed', 'Cancelled', 'Resolved', 'Test Run') group by week) c1;"
SQLite_with_weeks_to_Confluence_Chart("bar", "Core Coverage (weekly to date) with Automation", "", "Week Starting", "Percent of Functions Covered To Date", "horizontal", "false", "true", "800", "480", "false", "", "after", "1,2", "1,2", "vertical", "", "", "day", "", "", "true", "", "", "100", "5", "", $query)

; select week, count(week) from (select Key, strftime('%W', Date) as week, Status from SubTaskStateHistory group by Key, strftime('%W', Date) having SubTaskStateHistory.rowid = min(SubTaskStateHistory.rowid)) where Status in ('Waiting For Build', 'In Progress', 'Beta Testing', 'Ready', 'Closed', 'Cancelled', 'Resolved', 'Test Run') group by week;





_JiraShutdown()
_SQLite_Close()
_SQLite_Shutdown()

_Toast_Show(0, $app_name, "Uploading reports to confluence", -30, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", $confluence_ancestor_key, $confluence_page_key, $app_name, $storage_format)






Func SQLite_with_weeks_to_Confluence_Chart($type, $title, $subtitle, $xlabel, $ylabel, $orientation, $legend, $stacked, $width, $height, $show_shapes, $opacity, $data_display, $tables, $columns, $data_orientation, $time_series, $date_format, $time_period, $language, $country, $forgive, $colors, $range_axis_lower_bound, $range_axis_upper_bound, $range_axis_tick_unit, $category_label_position, $query)

	$storage_format = $storage_format & '<ac:structured-macro ac:name=\"chart\">' & @CRLF

	if StringLen($type) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"type\">' & $type & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($title) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"title\">' & $title & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($subtitle) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"subTitle\">' & $subtitle & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($xlabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"xLabel\">' & $xlabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($ylabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"yLabel\">' & $ylabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($orientation) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"orientation\">' & $orientation & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($legend) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"legend\">' & $legend & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($stacked) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"stacked\">' & $stacked & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($width) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"width\">' & $width & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($height) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"height\">' & $height & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($show_shapes) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"showShapes\">' & $show_shapes & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($opacity) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"opacity\">' & $opacity & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_display) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataDisplay\">' & $data_display & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($tables) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"tables\">' & $tables & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($columns) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"columns\">' & $columns & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_orientation) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataOrientation\">' & $data_orientation & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_series) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timeSeries\">' & $time_series & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($date_format) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dateFormat\">' & $date_format & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_period) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timePeriod\">' & $time_period & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($language) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"language\">' & $language & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($country) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"country\">' & $country & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($forgive) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"forgive\">' & $forgive & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($colors) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"colors\">' & $colors & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_lower_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisLowerBound\">' & $range_axis_lower_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_upper_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisUpperBound\">' & $range_axis_upper_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_tick_unit) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisTickUnit\">' & $range_axis_tick_unit & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($category_label_position) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"categoryLabelPosition\">' & $category_label_position & '</ac:parameter>' & @CRLF
	EndIf

	Local $double_quotes = """"

	if $confluence_html = true Then

		$double_quotes = "\"""
	EndIf

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

	;		$storage_format = $storage_format &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$storage_format = $storage_format &	"<ac:rich-text-body>" & @CRLF
			$storage_format = $storage_format &	"<table>" & @CRLF
			$storage_format = $storage_format &	"<tbody>" & @CRLF
			$storage_format = $storage_format & "<tr>" & @CRLF

			for $i = 0 to ($num_cols - 1)

				$storage_format = $storage_format & "<th>" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$storage_format = $storage_format & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$storage_format = $storage_format & "<tr>"

				for $j = 0 to ($num_cols - 1)

					if $j = 0 Then

						$aResult[$i][$j] = _DateFromWeekNumber(2019, ($aResult[$i][$j] + 1))
					Else

	;						$aResult[$i][$j] = StringReplace($aResult[$i][$j], " \</td>", " \\</td>")
						$aResult[$i][$j] = StringRegExpReplace($aResult[$i][$j], "([^\\])\\$", "$1\\\\")
	;						$a = StringRegExpReplace($a, "([^\\])\\$", "$1\\\\")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "<br>", "<br/>")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "&", "&amp;")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], """", "\""")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "\\""", "\""")
					EndIf

					$storage_format = $storage_format & "<td>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$storage_format = $storage_format & "</tr>" & @CRLF
			Next

			$storage_format = $storage_format &	"</tbody>" & @CRLF
			$storage_format = $storage_format &	"</table>" & @CRLF
			$storage_format = $storage_format &	"</ac:rich-text-body>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf

	$storage_format = $storage_format &	"</ac:structured-macro><br /><br />" & @CRLF

EndFunc









Func SQLite_to_Confluence_Chart($type, $title, $subtitle, $xlabel, $ylabel, $legend, $stacked, $width, $height, $show_shapes, $opacity, $data_display, $tables, $columns, $data_orientation, $time_series, $date_format, $time_period, $language, $country, $forgive, $colors, $range_axis_lower_bound, $range_axis_upper_bound, $category_label_position, $query)

	$storage_format = $storage_format & '<ac:structured-macro ac:name=\"chart\">' & @CRLF

	if StringLen($type) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"type\">' & $type & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($title) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"title\">' & $title & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($subtitle) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"subTitle\">' & $subtitle & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($xlabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"xLabel\">' & $xlabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($ylabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"yLabel\">' & $ylabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($legend) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"legend\">' & $legend & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($stacked) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"stacked\">' & $stacked & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($width) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"width\">' & $width & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($height) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"height\">' & $height & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($show_shapes) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"showShapes\">' & $show_shapes & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($opacity) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"opacity\">' & $opacity & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_display) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataDisplay\">' & $data_display & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($tables) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"tables\">' & $tables & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($columns) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"columns\">' & $columns & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_orientation) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataOrientation\">' & $data_orientation & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_series) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timeSeries\">' & $time_series & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($date_format) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dateFormat\">' & $date_format & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_period) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timePeriod\">' & $time_period & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($language) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"language\">' & $language & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($country) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"country\">' & $country & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($forgive) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"forgive\">' & $forgive & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($colors) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"colors\">' & $colors & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_lower_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisLowerBound\">' & $range_axis_lower_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_upper_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisUpperBound\">' & $range_axis_upper_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($category_label_position) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"categoryLabelPosition\">' & $category_label_position & '</ac:parameter>' & @CRLF
	EndIf

	Local $double_quotes = """"

	if $confluence_html = true Then

		$double_quotes = "\"""
	EndIf

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

	;		$storage_format = $storage_format &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$storage_format = $storage_format &	"<ac:rich-text-body>" & @CRLF
			$storage_format = $storage_format &	"<table>" & @CRLF
			$storage_format = $storage_format &	"<tbody>" & @CRLF
			$storage_format = $storage_format & "<tr>" & @CRLF

			for $i = 0 to ($num_cols - 1)

				$storage_format = $storage_format & "<th>" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$storage_format = $storage_format & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$storage_format = $storage_format & "<tr>"

				for $j = 0 to ($num_cols - 1)

;						$aResult[$i][$j] = StringReplace($aResult[$i][$j], " \</td>", " \\</td>")
					$aResult[$i][$j] = StringRegExpReplace($aResult[$i][$j], "([^\\])\\$", "$1\\\\")
;						$a = StringRegExpReplace($a, "([^\\])\\$", "$1\\\\")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], "<br>", "<br/>")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], "&", "&amp;")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], """", "\""")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], "\\""", "\""")
					$storage_format = $storage_format & "<td>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$storage_format = $storage_format & "</tr>" & @CRLF
			Next

			$storage_format = $storage_format &	"</tbody>" & @CRLF
			$storage_format = $storage_format &	"</table>" & @CRLF
			$storage_format = $storage_format &	"</ac:rich-text-body>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf

	$storage_format = $storage_format &	"</ac:structured-macro><br /><br />" & @CRLF

EndFunc


Func Update_Confluence_Page($url, $jira_username, $jira_password, $space_key, $ancestor_key, $page_key, $page_title, $page_body)

	_ConfluenceSetup()
	_ConfluenceDomainSet($url)
	_ConfluenceLogin($confluence_username, $confluence_decrypted_password)
	_ConfluenceUpdatePage($space_key, $ancestor_key, $page_key, $page_title, $page_body)
	_ConfluenceShutdown()

EndFunc


Func SQLite_to_Confluence_Chart_for_week($datetime)

	Local $datetime_part = StringSplit(StringLeft($datetime, 10), "/", 3)
	Local $month_name = _DateToMonth($datetime_part[1])

	Local $day1_date = StringLeft(_DateAdd("D", 0, $datetime), 10)
	Local $day2_date = StringLeft(_DateAdd("D", 1, $datetime), 10)
	Local $day3_date = StringLeft(_DateAdd("D", 2, $datetime), 10)
	Local $day4_date = StringLeft(_DateAdd("D", 3, $datetime), 10)
	Local $day5_date = StringLeft(_DateAdd("D", 4, $datetime), 10)
	Local $day6_date = StringLeft(_DateAdd("D", 5, $datetime), 10)
	Local $day7_date = StringLeft(_DateAdd("D", 6, $datetime), 10)

;	Local $query = "select aaa as 'Test Name', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day1_date & "'), 0) as '" & $day1_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day2_date & "'), 0) as '" & $day2_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day3_date & "'), 0) as '" & $day3_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day4_date & "'), 0) as '" & $day4_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day5_date & "'), 0) as '" & $day5_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day6_date & "'), 0) as '" & $day6_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day7_date & "'), 0) as '" & $day7_date & "' from (select distinct(Test) as aaa from SavedRuntime);"
	Local $query = "select strftime('%W', CreatedOn) as 'Week', sum(SavedRuntime) as 'Saved Runtime' from SavedRuntime group by week;"
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	SQLite_to_Confluence_Chart("bar", "Run Time Saved with Automation", "Week Starting " & Number($datetime_part[2]) & " " & $month_name & " " & Number($datetime_part[0]), "Run Date", "Minutes Saved", "true", "true", "800", "480", "false", "", "after", "1,2,3,4,5,6,7,8", "1,2,3,4,5,6,7,8", "horizontal", "", "", "day", "", "", "true", "", "", "", "", $query)
	SQLite_to_Confluence_Chart("bar", "Run Time Saved with Automation", "", "Week Number", "Minutes Saved", "true", "false", "800", "480", "false", "", "after", "1,2", "1,2", "vertical", "", "", "day", "", "", "true", "", "", "", "", $query)

EndFunc


; The week with the first Thursday of the year is week number 1.
Func _DateFromWeekNumber($iYear, $iWeekNum)
    Local $Date, $sFirstDate = _DateToDayOfWeek($iYear, 1, 1)
    If $sFirstDate < 6 Then
        $Date = _DateAdd("D", 2 - $sFirstDate, $iYear & "/01/01")
    ElseIf $sFirstDate = 6 Then
        $Date = _DateAdd("D", $sFirstDate - 3, $iYear & "/01/01")
    ElseIf $sFirstDate = 7 Then
        $Date = _DateAdd("D", $sFirstDate - 5, $iYear & "/01/01")
    EndIf
    ;ConsoleWrite(_DateToDayOfWeek($iYear, 1, 1) &"  ")
    Local $aDate = StringSplit($Date, "/", 2)
    Return _DateAdd("w", $iWeekNum - 1, $aDate[0] & "/" & $aDate[1] & "/" & $aDate[2])
EndFunc   ;==>_DateFromWeekNumber

