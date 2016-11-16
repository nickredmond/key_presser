Dim objResult

Const intervalKey = "INTERVAL"
Const durationKey = "DURATION"
Const unitKey = "UNIT"

Const millisecondsUnit = "MS"
Const secondsUnit = "S"
Const minutesUnit = "M"
Const hoursUnit = "H"

Const millisecondsCommonName = "milliseconds"
Const millisecondsMultiplier = 1
Const millisecondsDateAddInterval = "---"

interval = 0
duration = 0
unit = ""

Set args = WScript.Arguments
For Each arg In args
	tokens = Split(arg, "=")
	If StrComp(tokens(0), intervalKey, 1) = 0 Then
		interval = tokens(1)
	ElseIf StrComp(tokens(0), durationKey, 1) = 0 Then
		duration = tokens(1)
	ElseIf StrComp(tokens(0), unitKey, 1) = 0 Then
		unit = tokens(1)
	Else
		WScript.Echo("ERROR: Unsupported Argument '" & arg & "'")
	End If
Next

multiplier = 0
unitCommonName = 0
dateAddInterval = ""

If NOT (unit = "") Then
	If unit = millisecondsUnit Then
		multiplier = millisecondsMultiplier
		unitCommonName = millisecondsCommonName
		dateAddInterval = millisecondsDateAddInterval
	ElseIf unit = secondsUnit Then
		multiplier = 1000
		unitCommonName = "seconds"
		dateAddInterval = "s"
	ElseIf unit = minutesUnit Then
		multiplier = 60 * 1000
		unitCommonName = "minutes"
		dateAddInterval = "n"
	ElseIf unit = hoursUnit Then
		multiplier = 60 * 60 * 1000
		unitCommonName = "hours"
		dateAddInterval = "h"
	Else
		multiplier = millisecondsMultiplier
		unitCommonName = millisecondsCommonName
		dateAddInterval = millisecondsDateAddInterval
		WScript.Echo("ERROR: Unsupported unit '" & unit & "'; using the default unit '" & millisecondsUnit & " (" & millisecondsCommonName & ")', instead.")
	End If
End If

If unit = "" Then
	WScript.Echo("INFO: No unit specified; using the default unit '" & millisecondsUnit & " (" & millisecondsCommonName & ")', instead.")
End If

WScript.Echo("INFO: Registered to press NUMLOCK every " & interval & " " & unitCommonName & ", for the next " & duration & " " & unitCommonName)

maxIterations = duration / interval
iterationCount = 0
sleepDurationMillis = interval * multiplier
Set objShell = WScript.CreateObject("WScript.Shell")

dateAddNumber = 0
If dateAddInterval = millisecondsDateAddInterval Then
	dateAddNumber = interval / 1000
	dateAddInterval = "s"
Else
	dateAddNumber = interval
End If

Do While iterationCount < maxIterations
	objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
	WScript.Echo("-----------")
	WScript.Echo("INFO [" & FormatDateTime(Now()) & "]: Pressed NUMLOCK twice")
	WScript.Echo("-- Next NUMLOCK press at " & DateAdd(dateAddInterval, dateAddNumber, Now()))

	iterationCount = iterationCount + 1
	remainingIterations = maxIterations - iterationCount
	WScript.Echo("-- " & remainingIterations & " iterations remain")
	If remainingIterations > 0 Then
		WScript.Sleep (sleepDurationMillis)
	End If
Loop