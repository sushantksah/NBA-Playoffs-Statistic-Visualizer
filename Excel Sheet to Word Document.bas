Attribute VB_Name = "Module2"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID:  169060628
' Date: 03/31/2025
' Program title: Assignment 5 Word Report
' Description: Excel to Word
'===========================================================+

Sub writeInDoc()
    Dim ws As Worksheet
    
    ' Late Binding makes more sense to use with this use-case
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdSel As Object
    Dim filePath As String
    Dim i As Integer
    
    ' Didn't know how to paste multiple charts so I used the microsoft learn page
    ' https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject
    ' Found out you can use chartObjects in order to iterate through charts
    Dim chartObj As ChartObject

    Set ws = ThisWorkbook.Sheets("Histogram Data")

    ' Start Word application
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    ' Create new document
    Set wdDoc = wdApp.Documents.Add
    Set wdSel = wdApp.Selection
    
    ' Add the text from the corresponding cell into the word report
    With wdSel
        .TypeText Text:=ws.Range("A1").Value
        .TypeParagraph
        
        .TypeText Text:=ws.Range("A2").Value & " "
        .TypeParagraph
        
        .TypeText Text:=ws.Range("A3").Value & " "
        .TypeText Text:=ws.Range("A4").Value
        .TypeParagraph
        
        .TypeText Text:=ws.Range("A6").Value & " "
        .TypeText Text:=ws.Range("A7").Value & " "
        .TypeText Text:=ws.Range("A8").Value
        .TypeParagraph
        .TypeParagraph
        
        .TypeText Text:=ws.Range("A17").Value
    End With
    
    ' Add the 6 charts for all of the major stat categories
    For i = 1 To 6
        Set chartObj = ws.ChartObjects(i)
        
        ' Copy the image
        chartObj.Chart.ChartArea.Copy
        
        ' Paste it into the file
        With wdSel
            .Paste
            .TypeParagraph
        End With
    Next i
    
    ' Finish adding text
    With wdSel
        .TypeText Text:=ws.Range("A12").Value
        .TypeParagraph
        .TypeText Text:=ws.Range("A13").Value
        .TypeParagraph
        .TypeText Text:=ws.Range("A14").Value
    End With
    
    ' Save document
    filePath = ThisWorkbook.Path & "\Playoffs Statistic Visualizer Report.docx"
    wdDoc.SaveAs2 filePath

    ' Close Word
    wdDoc.Close

    ' Clean Up
    Set wdSel = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing

    MsgBox "Report created!: " & filePath, vbInformation

End Sub
