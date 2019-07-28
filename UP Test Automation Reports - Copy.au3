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

Global $app_name = "UP Test Automation Runs Reports"
Global $confluence_page_key = "469172334"
Global $confluence_ancestor_key = "469172330"
Global $testrail_project_id = 49
Global $github_release_url = "https://github.com/seanhaydongriffin/UP-Test-Automation-Runs-Reports/releases/download/v0.1/UP.Test.Automation.Runs.Reports.portable.exe"

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

_JiraShutdown()


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

	if StringInStr($plan[$i][1], "UP ") > 0 And StringInStr($plan[$i][1], " SIT") > 0 Then

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
Local $previous_quarter_date = StringLeft(_DateAdd("D", $number_of_days_to_last_monday - (7 * 11), _NowCalc()), 10)
$created_after_date = _DateDiff( 's',"1970/01/01 00:00:00", $previous_quarter_date & " 00:00:00")

Local $tomorrow_date = StringLeft(_DateAdd("D", 1, _NowCalc()), 10)
$created_before_date = _DateDiff( 's',"1970/01/01 00:00:00", $tomorrow_date & " 00:00:00")

Local $saved_runtime = _TestRailGetCreatedOnTestSavedRuntimeForRuns($testrail_project_id, $run_suite, $created_after_date, $created_before_date, $testcase_complexity_execution_weightage_dict)
;Local $saved_runtime = _TestRailGetCreatedOnTestSavedRuntimeForRuns($testrail_project_id, $run_suite, "", "", $testcase_complexity_execution_weightage_dict)

for $i = 0 to (UBound($saved_runtime) - 1)

	$saved_runtime[$i][0] = StringLeft($saved_runtime[$i][0], 10)

	$query = "INSERT INTO SavedRuntime (CreatedOn,Test,SavedRuntime) VALUES ('" & $saved_runtime[$i][0] & "','" & $saved_runtime[$i][1] & "'," & $saved_runtime[$i][2] & ");"
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	_SQLite_Exec(-1, $query) ; INSERT Data
Next

Local $todays_day_of_week = _DateToDayOfWeek(@YEAR, @MON, @MDAY)
Local $number_of_days_to_last_monday = 0

if $todays_day_of_week = 1 Then

	$number_of_days_to_last_monday = -6
Else

	$number_of_days_to_last_monday = 2 - $todays_day_of_week
EndIf

_Toast_Show(0, $app_name, "creating report", -30, False, True)

$storage_format = $storage_format & '<a href=\"' & $github_release_url & '\">Click to update this page</a><br /><br />'

Local $start_datetime_of_week = _DateAdd("D", $number_of_days_to_last_monday, _NowCalc())
SQLite_to_Confluence_Chart_for_week($start_datetime_of_week)

for $i = 1 to 11

	$start_datetime_of_week = _DateAdd("D", -7, $start_datetime_of_week)
	SQLite_to_Confluence_Chart_for_week($start_datetime_of_week)
Next

_SQLite_Close()
_SQLite_Shutdown()

_Toast_Show(0, $app_name, "uploading report to confluence", -30, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", $confluence_ancestor_key, $confluence_page_key, $app_name, $storage_format)




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

	Local $query = "select aaa as 'Test Name', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day1_date & "'), 0) as '" & $day1_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day2_date & "'), 0) as '" & $day2_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day3_date & "'), 0) as '" & $day3_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day4_date & "'), 0) as '" & $day4_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day5_date & "'), 0) as '" & $day5_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day6_date & "'), 0) as '" & $day6_date & "', ifnull((select sum(SavedRuntime) from SavedRuntime where Test = aaa and CreatedOn = '" & $day7_date & "'), 0) as '" & $day7_date & "' from (select distinct(Test) as aaa from SavedRuntime);"
	SQLite_to_Confluence_Chart("bar", "Run Time Saved with Automation", "Week Starting " & Number($datetime_part[2]) & " " & $month_name & " " & Number($datetime_part[0]), "Run Date", "Minutes Saved", "true", "true", "640", "480", "false", "", "after", "1,2,3,4,5,6,7,8", "1,2,3,4,5,6,7,8", "horizontal", "", "", "day", "", "", "true", "", "", "", "", $query)

EndFunc