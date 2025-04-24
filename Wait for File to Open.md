```
Dim tries As Integer
Const MaxTries As Integer = 10
tries = 0
found = False

Do
    ' Check for workbook first
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, exportWbName, vbTextCompare) = 0 Then
            Set exportWb = wb
            found = True
            Exit Do
        End If
    Next wb
    
    ' If workbook not found, and maximum tries reached, exit loop
    If tries >= MaxTries Then Exit Do
    
    ' Yield control to process pending events and then wait
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    tries = tries + 1
Loop

If Not found Then
    ' Handle error if workbook did not open in time
    GoTo ErrHandler
End If
```