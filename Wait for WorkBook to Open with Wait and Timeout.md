```
startTime = Timer

Do
	found = False
	
	' Loop through all open workbooks
	For Each wb In Application.Workbooks
		If UCase(wb.Name) = UCase(exportWbName) Then
			Set exportWb = wb
			found = True
			Exit Do  ' Exit the loop immediately if the workbook is found
		End If
	Next wb
	
	' Check if 60 seconds have elapsed
	If Timer - startTime >= 60 Then
		GoTo ErrorHandler
	End If
	
	DoEvents  ' Yield control to allow other events to be processed
	
	' Pause for 1 second to give the external process time to open the workbook
    Application.Wait Now + TimeValue("00:00:01")
Loop
```