#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\OneDrive\Pictures\TimeTracker.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.15.0 (Beta)
	Author:         myName

	Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <Excel.au3>
#include <GuiListView.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <Constants.au3>
#include <Misc.au3>
Global $count = 0
$Form1 = GUICreate("Time Tracker", 405, 429, 192, 124)
$Label1 = GUICtrlCreateLabel("Time Tracker", 142, 8, 120, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("    00:00:00", 72, 40, 244, 84)
GUICtrlSetFont(-1, 50, 400, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("Project/ticket working on", 87, 90, 214, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$Input1 = GUICtrlCreateInput("", 96, 130, 201, 21)
$Label3 = GUICtrlCreateLabel("All time entries", 120, 162, 130, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$ListView1 = GUICtrlCreateListView("Task|Start Time|End time|Total Time", 8, 194, 385, 175)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 73)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 73)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 73)
$Button1 = GUICtrlCreateButton("Start Timer", 24, 376, 89, 41, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Button2 = GUICtrlCreateButton("Export to file", 152, 376, 89, 41, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Button3 = GUICtrlCreateButton("Cancel", 280, 376, 81, 41, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$dummy = GUICtrlCreateDummy ()
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


While 1
	$x = Int(Stopwatch() / 100)
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			If _GUICtrlListView_GetItemCount ( $ListView1 ) > 0 Then
				If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
				$iMsgBoxAnswer = MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_SYSTEMMODAL,"Save info?","Do you want to save your time entries?")
				Select
					Case $iMsgBoxAnswer = $IDYES
						$exfile2 = FileSaveDialog ( "Name of exported document.", "", "Excel document (*.xlsx)", 18, "", $Form1 )
						If $exfile2 = "" Then

						Else
							$num2 = _GUICtrlListView_GetItemCount ( $ListView1 )
							Local $array2[$num2 + 1][4]
							$array2[0][0] = "Task"
							$array2[0][1] = "Start Time"
							$array2[0][2] = "End Time"
							$array2[0][3] = "Total Time"
							For $y = 0 To $num2 - 1 Step 1
								$listarray2 = _GUICtrlListView_GetItemTextArray ( $ListView1, $y )
								$array2[$y + 1][0] = $listarray2[1]
								$array2[$y + 1][1] = $listarray2[2]
								$array2[$y + 1][2] = $listarray2[3]
								$array2[$y + 1][3] = $listarray2[4]
							Next
							$oExcel2 = _Excel_Open ()
							$oWorksheet2 = _Excel_BookNew ( $oExcel2 )
							_Excel_RangeWrite ( $oWorksheet2, Default, $array2 )
							_Excel_BookSaveAs ( $oWorksheet2, $exfile2 )
							_Excel_Close ( $oExcel2 )
							Exit
						EndIf
					Case $iMsgBoxAnswer = $IDNO
						Exit
				EndSelect
			Else
				Exit
			EndIf
		Case $Form1
			ToolTip ("")
		Case $Input1
			If _IsPressed ( "0D" ) Then
				ToolTip ("")
				Stopwatch()
				If @extended = 1 Then
					If GUICtrlRead($Input1) = "" Then
						$sToolTipAnswer = ToolTip("You need to specify the task in order to catalog the time.", Default, Default, "Enter task")
					Else
						Starttime(@HOUR, @MIN)
						GUICtrlSetData ( $Button1, "Stop Timer" )
						GUICtrlSetState ( $Button2, $GUI_DISABLE )
						GUICtrlSetData ( $Input1, "" )
						GUICtrlSetState ( $Input1, $GUI_DISABLE )
						GUICtrlSetState ( $Button3, $GUI_DISABLE )
						Stopwatch(1)
					EndIf
				EndIf
			EndIf

		Case $Button1
			ToolTip ("")
			Stopwatch()
			$store = @extended
			If $store = 1 Then
				If GUICtrlRead($Input1) = "" Then
					$sToolTipAnswer = ToolTip("You need to specify the task in order to catalog the time.", Default, Default, "Enter task")
				Else
					Starttime(@HOUR, @MIN)
					GUICtrlSetData ( $Button1, "Stop Timer" )
					GUICtrlSetState ( $Button2, $GUI_DISABLE )
					GUICtrlSetData ( $Input1, "" )
					GUICtrlSetState ( $Input1, $GUI_DISABLE )
					GUICtrlSetState ( $Button3, $GUI_DISABLE )
					Stopwatch ( 1 )
				EndIf
			Else
				Endtime ( @HOUR, @MIN )
				GUICtrlSetData ( $Button1, "Start Timer" )
				GUICtrlSetState ( $Button2, $GUI_ENABLE )
				GUICtrlSetState ( $Button3, $GUI_ENABLE )
				GUICtrlSetState ( $Input1, $GUI_ENABLE )
				Stopwatch ( 2 )
			EndIf


		Case $Button2
			$exfile = FileSaveDialog ( "Name of exported document.", "", "PDF file (*.pdf)", 18, "", $Form1 )
			If $exfile = "" Then

			Else
				$num = _GUICtrlListView_GetItemCount ( $ListView1 )
				Local $array[$num + 1][4]
				$array[0][0] = "Task"
				$array[0][1] = "Start Time"
				$array[0][2] = "End Time"
				$array[0][3] = "Total Time"
				For $y = 0 To $num - 1 Step 1
					$listarray = _GUICtrlListView_GetItemTextArray ( $ListView1, $y )
					$array[$y + 1][0] = $listarray[1]
					$array[$y + 1][1] = $listarray[2]
					$array[$y + 1][2] = $listarray[3]
					$array[$y + 1][3] = $listarray[4]
				Next
				$oExcel = _Excel_Open ()
				$oWorksheet = _Excel_BookNew ( $oExcel )
				_Excel_RangeWrite ( $oWorksheet, Default, $array )
				_Excel_Export ( $oExcel, $oWorksheet, $exfile, Default, Default, Default, Default, Default, True )
				_Excel_BookSaveAs ( $oWorksheet, _GetFilenameDrive ( $exfile ) & _GetFilenamePath ( $exfile ) & _GetFilename ( $exfile ) & ".xlsx" )
				_Excel_Close ( $oExcel )
			EndIf

		Case $Button3
			If _GUICtrlListView_GetItemCount ( $ListView1 ) > 0 Then
				If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
				$iMsgBoxAnswer = MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_SYSTEMMODAL,"Save info?","Do you want to save your time entries?")
				Select
					Case $iMsgBoxAnswer = $IDYES
						$exfile2 = FileSaveDialog ( "Name of exported document.", "", "Excel document (*.xlsx)", 18, "", $Form1 )
						If $exfile2 = "" Then

						Else
							$num2 = _GUICtrlListView_GetItemCount ( $ListView1 )
							Local $array2[$num2 + 1][4]
							$array2[0][0] = "Task"
							$array2[0][1] = "Start Time"
							$array2[0][2] = "End Time"
							$array2[0][3] = "Total Time"
							For $y = 0 To $num2 - 1 Step 1
								$listarray2 = _GUICtrlListView_GetItemTextArray ( $ListView1, $y )
								$array2[$y + 1][0] = $listarray2[1]
								$array2[$y + 1][1] = $listarray2[2]
								$array2[$y + 1][2] = $listarray2[3]
								$array2[$y + 1][3] = $listarray2[4]
							Next
							$oExcel2 = _Excel_Open ()
							$oWorksheet2 = _Excel_BookNew ( $oExcel2 )
							_Excel_RangeWrite ( $oWorksheet2, Default, $array2 )
							_Excel_BookSaveAs ( $oWorksheet2, $exfile2 )
							_Excel_Close ( $oExcel2 )
							Exit
						EndIf
					Case $iMsgBoxAnswer = $IDNO
						Exit
				EndSelect
			Else
				Exit
			EndIf
	EndSwitch
			If $x <> Int(Stopwatch() / 100) Then
				$totsec = Int(Stopwatch() / 1000) ; ms to sec
				$hr = Int($totsec / 3600) ; hours
				$mn = Int(($totsec - ($hr * 3600)) / 60) ; minutes
				$sc = Int(($totsec - ($hr * 3600) - ($mn * 60))) ; seconds
				$tn = Int((Int(Stopwatch() / 100) - ($hr * 36000) - ($mn * 600) - ($sc * 10))) ; tenths of a second
				GUICtrlSetData($Label4, "    " & StringFormat("%02s", $hr) & ":" & StringFormat("%02s", $mn) & ":" & StringFormat("%02s", $sc) & "." & StringFormat("%01s", $tn))
				If $mn >= 15 Then
					GUICtrlSetColor ( $Label4, 0xFF0000 )
				Else
					GUICtrlSetColor ( $Label4, 0x000000 )
				EndIf
			EndIf
		WEnd
		Func Stopwatch($ToggleTo = 4)
			Static Local $Paused = True
			Static Local $Stopwatch = 0
			Static Local $TotalTime = 0
			Switch $ToggleTo
				Case 0 ; pause counter
					If $Paused Then
						SetExtended($Paused) ; $Paused status
						Return $TotalTime ; already paused, just return current $TotalTime
					Else
						$TotalTime += TimerDiff($Stopwatch)
						$Paused = True
						SetExtended($Paused)
						Return $TotalTime
					EndIf
				Case 1 ; unpause counter
					If $Paused Then
						$Stopwatch = TimerInit()
						$Paused = False
						SetExtended($Paused)
						Return $TotalTime
					Else
						SetExtended($Paused)
						Return $TotalTime + TimerDiff($Stopwatch)
					EndIf
				Case 2 ; reset to 0 and pause
					$Paused = True
					$TotalTime = 0
					SetExtended($Paused)
					Return $TotalTime
				Case 3 ; reset to 0 and restart
					$Paused = False
					$TotalTime = 0
					$Stopwatch = TimerInit()
					SetExtended($Paused)
					Return $TotalTime
				Case 4 ; return status
					SetExtended($Paused)
					If $Paused Then
						Return $TotalTime
					Else
						Return $TotalTime + TimerDiff($Stopwatch)
					EndIf
			EndSwitch
		EndFunc   ;==>Stopwatch
		Func Starttime($hour, $min)
			$pm = False
			$tm = ""
			If Int($hour) >= 12 Then
				$pm = True
				$tm = " PM"
				If Int($hour) > 12 Then
					$hour = Int($hour) - 12
				EndIf

			Else
				$pm = False
				$tm = " AM"
				If Int($hour) = 0 Then
					$hour = Int($hour) + 12
				EndIf

			EndIf
			GUICtrlCreateListViewItem(GUICtrlRead($Input1) & "|" & $hour & ":" & $min & $tm, $ListView1)
		EndFunc   ;==>Starttime
		Func Endtime($hour2, $min2)
			$pm2 = False
			$tm2 = ""
			If Int($hour2) >= 12 Then
				$pm2 = True
				$tm2 = " PM"
				If Int($hour2) > 12 Then
					$hour2 = Int($hour2) - 12
				EndIf

			Else
				$pm2 = False
				$tm2 = " AM"
				If Int($hour2) = 0 Then
					$hour2 = Int($hour2) + 12
				EndIf

			EndIf
			$time = Int ( Round ( Stopwatch () ) ) / 1000
			$hr2 = Int($time / 3600) ; hours
			$mn2 = Int(($time - ($hr2 * 3600)) / 60) ; minutes
			$sc2 = Int(($time - ($hr2 * 3600) - ($mn2 * 60))) ; seconds
			_GUICtrlListView_AddSubItem($ListView1, $count, $hour2 & ":" & $min2 & $tm2, 2)
			_GUICtrlListView_AddSubItem($ListView1, $count, StringFormat("%02s", $hr2) & ":" & StringFormat("%02s", $mn2) & ":" & StringFormat("%02s", $sc2), 3)
			$count += 1
		EndFunc   ;==>Endtime
Func _GetFilename($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return $oObjectFile.FileName
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilename
Func _GetFilenameDrive($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return StringUpper($oObjectFile.Drive)
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenameDrive

Func _GetFilenamePath($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return $oObjectFile.Path
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenamePath
