' This is useful when you need multiple maps in the same Sub or Function since each dictionary is a variable 
```
Sub MainSub()

	Dim colMap As Object
	
	' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    colMap.Add "ID", 1
	colMap.Add "Cliente", 2
    colMap.Add "PM", 3
    colMap.Add "PEP", 4
    colMap.Add "Tipo", 5
    colMap.Add "Valor Total", 6
    colMap.Add "Custo", 7
    colMap.Add "Apolice", 8
    colMap.Add "Percentual", 9
    colMap.Add "Inicio Vigencia", 10
    colMap.Add "Fim Vigencia", 11
    colMap.Add "Status", 12
    
	Dim headers As Object
    
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
	
	 --- Start of main code ---
	 Some code here
	 --- End of main code ---
End Sub
```