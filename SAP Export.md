```
Function UpdateCJI3(wb As Workbook, PEP As String) As Boolean

Dim temp As Double
temp = Timer
Debug.Print "UpdateCJI3 Start"

    Dim exportWb As Workbook
    Dim Workbook As Workbook
    Dim wsCJI3 As Worksheet
    Dim exportWs As Worksheet
    Dim exportWbName As String
    Dim exportWbPath As String
    Dim EndDate As String
    Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    
    On Error Resume Next
    Set wsCJI3 = wb.Sheets("CJI3")
    On Error GoTo 0
    
    ' Check if the "CJI3" sheet exists
    If wsCJI3 Is Nothing Then
        UpdateCJI3 = False
        Exit Function
    End If
    
    ' Name of the workbook to find
    exportWbName = "CJI3-" & PEP
    
    ' Set end date
    EndDate = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy")
    
    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count

Debug.Print "Setup time: " & Timer - temp
temp = Timer

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncji3"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
    session.findById("wnd[1]").sendVKey 0
    On Error GoTo 0
    
    ' Clear other fields
    session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_PROJN-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_NETNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_ACTVT-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_MATNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtR_KSTAR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtR_KSTAR-HIGH").Text = ""
    
    ' Search PEP
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = PEP
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").Text = "01.11.2000"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").Text = EndDate
    session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/CUSTO_CIDIO"
    session.findById("wnd[0]/usr/ctxtP_DISVAR").SetFocus
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo 0
    
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

    ' Wait for a new workbook to appear
    Do
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop
    
    ' Validate if the workbook was opened successfully
    If exportWb Is Nothing Then
        wsCJI3.UsedRange.ClearContents
        UpdateCJI3 = False
        Exit Function
    End If
    
    Set exportWs = exportWb.Sheets(1)
    
Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer
    
    ' Clear, copy and paste data from exportWs to wsCJI3
    If wsCJI3.AutoFilterMode Then wsCJI3.AutoFilter.ShowAllData ' Clear any applied filters
    wsCJI3.UsedRange.ClearContents
    exportWs.UsedRange.Copy
    wsCJI3.UsedRange.PasteSpecial
    
    ' Ensure columns A and B are converted to numbers
    With wsCJI3
        .Columns("A:B").NumberFormat = "0"  ' Set format to number
        .Columns("A:B").Value = .Columns("A:B").Value  ' Convert text to numbers
    End With
    
    ' Cleanup
    Application.CutCopyMode = False
    exportWb.Close False  ' Close the exported workbook without saving

Debug.Print "Project Review CJI3 Sheet update: " & Timer - temp
temp = Timer

    UpdateComentarios wb, wsCJI3
    
Debug.Print "Project Review comments update: " & Timer - temp
temp = Timer
    
    UpdateCJI3 = True
    
End Function
```