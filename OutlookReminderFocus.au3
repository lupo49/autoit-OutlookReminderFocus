; Name: OutlookReminderFocus.au3
;
; Author: M. Schulte - 2013-06-18 - Created
;
; Moves the reminder windows of Outlook 2010 in the foreground,
; when the Outlook COM fires the event ReminderFire.
;
; http://www.autoitscript.com/autoit3/docs/intro/ComRef.htm
; http://msdn.microsoft.com/en-us/library/office/aa171361%28v=office.11%29.aspx
; http://www.autoitscript.com/forum/topic/26745-outlook-com-and-events/
; http://www.autoitscript.com/forum/topic/26745-outlook-com-and-events/
; http://www.dimastr.com/outspy/home.htm

#AutoIt3Wrapper_Icon=OutlookReminderFocus.ico
#AutoIt3Wrapper_OutFile=OutlookReminderFocus.exe
#AutoIt3Wrapper_Res_Description=Moves Outlooks reminder windows on top, when a reminder is triggered.

; http://www.autoitscript.com/autoit3/docs/functions/AutoItSetOption.htm#WinTitleMatchMode
Opt("WinTitleMatchMode", 2)
; Set to 1 to turn on debug messages
$debug = 0;
$oEvtReminders = ""
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc") ; Initialize a COM error handler

; Prevent this application from starting Outlook
While 1
	If ProcessExists("outlook.exe") Then
		; Outlook has been started, exit loop and watch for ReminderFire events
		$outlook = ObjGet("", "Outlook.Application")

		If @error = 1 Then
			If $debug Then ConsoleWrite("Error getting obj ref to Outlook" & @CR)
			Exit
		EndIf

		If IsObj($outlook) Then
			$reminders 		= $outlook.Reminders;
			$oEvtReminders 	= ObjEvent($reminders, "OutlookEventReminders_");

			; Determine UI language of Outlook
			; Parameter: 	1 -> Installation language
			; 				2 -> User language
			; http://msdn.microsoft.com/de-de/library/office/ff862542.aspx
			$langID = $outlook.LanguageSettings.LanguageID(2)

			If $debug Then ConsoleWrite("Version is " & $outlook.Version & @CR)
			If $debug Then ConsoleWrite("Language is " & $langID & @CR)
			If $debug Then ConsoleWrite("Found outlook..." & @CR)

			If Not IsDeclared ("reminders") then
				MsgBox(0,"", "OutlookReminderFocus.exe - $event is NOT declared" & @CR)
				Exit
			EndIf

		EndIf
	EndIf

	; Remove handle of Outlook.Application, to allow a complete shutdown of Outlook
	If Not ProcessExists("outlook.exe") Then
		If $debug Then ConsoleWrite("Outlook process not found..." & @CR)
		If IsObj($oEvtReminders) Then
			If $debug Then ConsoleWrite("Outlook object released..." & @CR)
			$oEvtReminders = ""
		EndIf
	EndIf

	Sleep(5000)
WEnd

Func OutlookEventReminders_ReminderFire($obj)
	; Function will be executed, when Outlook triggers the reminder window
	Switch $langID
		Case 1031
			; Get window handle of reminder window (German outlook)
			If $debug Then ConsoleWrite("Get handle of German reminder window." & @CR)
			$rmHandle = WinGetHandle("[TITLE:Erinnerung; CLASS:#32770]", "")
		Case 1033
			If $debug Then ConsoleWrite("Get handle of English reminder window." & @CR)
			$rmHandle = WinGetHandle("[TITLE:Reminder; CLASS:#32770]", "")
	EndSwitch

	If @error Then
		If $debug Then ConsoleWrite("Could not find the correct window" & @CR)
	Else
		; Bring reminder windows to foreground
		If $debug Then ConsoleWrite("Found reminder window. Set window on top." & @CR)
		WinSetState($rmHandle, "", @SW_RESTORE)
		WinSetOnTop($rmHandle, "", 1)
	EndIf
EndFunc

Func MyErrFunc($oMyError)
    ConsoleWrite("COM error" & @CR)
    $HexNumber = hex($oMyError.number, 8)
    ConsoleWrite("COM Error: err.description is: " & @TAB & $oMyError.description & " err.number is: " & @TAB & $HexNumber & _
				  " err.scriptline is: " & @TAB & $oMyError.scriptline & @CR)
	
	Msgbox(0, "AutoItCOM Test","We intercepted a COM Error!"    		  & @CRLF  & @CRLF & _
             "err.description is: " 	& @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"   	& @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       	& @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   	& @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "  	& @TAB & $oMyError.scriptline     & @CRLF & _
             "err.source is: "     		& @TAB & $oMyError.source         & @CRLF & _
             "err.helpfile is: "       	& @TAB & $oMyError.helpfile       & @CRLF & _
             "err.helpcontext is: " 	& @TAB & $oMyError.helpcontext _
            )
    SetError(1)
EndFunc

; eof