```
Sub MainSub()
	dim colDict As Object
	Set colDict = CreateObject("Scripting.Dictionary")
	
	 --- Start of main code ---
	 Some code here
	 --- End of main code ---
End Sub
```

```
Public Function GetColumnHeadersMapping() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "ID", 1
    headers.Add "Cliente", 2
    headers.Add "PM", 3
    headers.Add "PEP", 4
    headers.Add "Tipo", 5
    headers.Add "Valor Total", 6
    headers.Add "Custo", 7
    headers.Add "Apolice", 8
    headers.Add "Percentual", 9
    headers.Add "Inicio Vigencia", 10
    headers.Add "Fim Vigencia", 11
    headers.Add "Status", 12
    
    Set GetColumnHeadersMapping = headers
End Function
```