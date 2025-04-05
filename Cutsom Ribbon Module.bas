Attribute VB_Name = "Module1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID:  169060628
' Date: 03/31/2025
' Program title: Assignment 5 Ribbon Module
' Description: Ribbon Buttons / Functionality
'===========================================================+

'Callback for btnDisplay onAction
' Change the Icon when you can
Sub displayUserForm(control As IRibbonControl)
    UserForm1.Show
End Sub

'Callback for btnCleary onAction
Sub ClearResults(control As IRibbonControl)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Filter Summary")
    On Error GoTo 0
 
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        
        MsgBox "Worksheet has been deleted!", vbExclamation
    Else
        MsgBox "The worksheet you are trying to delete does not exist!", vbCritical
    
    End If
    
End Sub

'Callback for btnGenerateDoc onAction
Sub writeInWord(control As IRibbonControl)
    Call writeInDoc
End Sub



