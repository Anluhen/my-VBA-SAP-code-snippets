```
Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public session As Object

Sub ExempleSub() 
	' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo CleanExit  ' Exit the function or sub
        End If
    Loop
Exit Sub

Function SetupSAPScripting() As Boolean
    
    ' Create the SAP GUI scripting engine object
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    
    If Not IsObject(SapGuiAuto) Or SapGuiAuto Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    On Error Resume Next
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    On Error GoTo 0
    
    If Not IsObject(SAPApplication) Or SAPApplication Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    ' Get the first connection and session
    On Error GoTo ErrorHandler
    Set Connection = SAPApplication.Children(0)
    Set session = Connection.Children(0)
    On Error GoTo 0
    
    SetupSAPScripting = True
    
    If False Then
ErrorHandler:
    SetupSAPScripting = False
    End If
    
End Function

Function EndSAPScripting()
    ' Clean up
    Set session = Nothing
    Set Connection = Nothing
    Set SAPApplication = Nothing
    Set SapGuiAuto = Nothing
End Function
```