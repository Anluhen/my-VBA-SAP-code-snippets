```
Public colDict As Object ' Dictionary to store column names and indexes

Sub GetAllColumnIndexes(ws As Worksheet, Optional ShowOnMacroList As Boolean = False)
    Set colDict = CreateObject("Scripting.Dictionary")
    
    ' Map internal aliases to actual header names to avoid maintenace
    dict.Add "Date", GetColumnIndex(ws, "Data Entrada no Relatório")
    dict.Add "PEP", GetColumnIndex(ws, "PEP")
    dict.Add "Market", GetColumnIndex(ws, "Mercado")
    dict.Add "Client", GetColumnIndex(ws, "CLIENTE")
    dict.Add "OV", GetColumnIndex(ws, "OV")
    dict.Add "ZVA1", GetColumnIndex(ws, "ZVA1")
    dict.Add "ZETO", GetColumnIndex(ws, "ZETO")
    dict.Add "PaymentTerms", GetColumnIndex(ws, "Cond Pgto")
    dict.Add "OrderLocation", GetColumnIndex(ws, "Local/ Pedido")
    dict.Add "Incoterm", GetColumnIndex(ws, "Incoterm")
    dict.Add "Incoterm2", GetColumnIndex(ws, "Incoterm 2")
    dict.Add "PM", GetColumnIndex(ws, "PM")
    dict.Add "Amount", GetColumnIndex(ws, "R$")
    dict.Add "BillingResp", GetColumnIndex(ws, "Resp.Fat.")
    dict.Add "BillingForecast", GetColumnIndex(ws, "PREV. FAT")
    dict.Add "StockStatus", GetColumnIndex(ws, "Situação Estoque")
    dict.Add "Checklist", GetColumnIndex(ws, "CheckList")
    dict.Add "Freight", GetColumnIndex(ws, "Frete/LIS")
    dict.Add "Status", GetColumnIndex(ws, "Situação")
    dict.Add "PhysicalStock", GetColumnIndex(ws, "Estoque físico atual")
    dict.Add "Notes", GetColumnIndex(ws, "Observações")
    
End Sub

Function GetColumnIndex(ws As Worksheet, headerName As String, Optional headerRow As Long = 1) As Long
    Dim col As Range
    For Each col In ws.Rows(headerRow).Cells
        If Trim(UCase(col.Value)) = Trim(UCase(headerName)) Then
            GetColumnIndex = col.Column
            Exit Function
        End If
    Next col
    GetColumnIndex = 0 ' Not found
End Function
```