VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Statistic Filtering"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "UserForm - Dynamic Query Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID:  169060628
' Date: 03/31/2025
' Program title: Assignment 5 UserForm Code
' Description: 2024 NBA Playoff Statistics
'===========================================================+

' Populate the UserForm with the defulat data and options defined below
Private Sub UserForm_Initialize()
    InitialStates
End Sub


Sub InitialStates()
    ' Adding the option to filter based on poisiton, in order of what the poistion appears at on
    ' basketball position charts to make it easier for the user to choose.
    ' Adding the option to choose none, which will show data for all poisitions if chosen.
    With positionListBox
        .AddItem "None"
        .AddItem "PG"
        .AddItem "SG"
        .AddItem "SF"
        .AddItem "PF"
        .AddItem "C"
        .ListIndex = 0
        
        ' To highlight Multiple options
        .MultiSelect = 1
    End With
    
    ' Add the option to choose team as a parameter, in alphebetical order to make it easier
    ' find the team user is looking for. As well as the option to choose no team, which will
    ' show the data for all teams if chosen.
    With teamListBox
        .AddItem "None"
        .AddItem "BOS"
        .AddItem "CLE"
        .AddItem "DAL"
        .AddItem "DEN"
        .AddItem "IND"
        .AddItem "LAL"
        .AddItem "LAC"
        .AddItem "MIN"
        .AddItem "MIA"
        .AddItem "MIL"
        .AddItem "NOP"
        .AddItem "NYK"
        .AddItem "OKC"
        .AddItem "ORL"
        .AddItem "PHI"
        .AddItem "PHO"
        .ListIndex = 0
        
        ' To highlight multiple options
        .MultiSelect = 1
    End With
    
    ' Default ages for the minimuma age selection
    minAgeTextBox.Value = 18
    maxAgeTextBox.Value = 40
    
    
    End Sub
    
    ' Cancel Button for UserForm
    Private Sub cancelButton_Click()
        Unload Me
    End Sub
    
    ' Clear Button
    Private Sub clearButton_Click()
        Dim ws As Worksheet
        
        ' Error handling to allow for ws to attemp to be set in case worksheet doesn't exist in the file
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Filter Summary")
        On Error GoTo 0
     
        ' Check if the worksheet exists
        If Not ws Is Nothing Then
            ' Turning off alerts to delete with no message
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            
            MsgBox "Worksheet has been deleted!", vbExclamation
        Else
            MsgBox "The worksheet you are trying to delete does not exist!", vbCritical
        
        End If
        
    End Sub
    
    ' Opening File to Select Database
    ' Didn't want users to have issues with the database not working
    ' Users inputting the database helps to avoid this.
    Private Sub browseButton_Click()
        Dim fd As FileDialog
        Dim notCancel As Boolean
        
        'msoFileDialogFilePicker to pick the file from a users file explorer
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        ' Prompting to open database from user's file explorer
        With fd
            notCancel = .Show
            If notCancel Then
                'Setting database name
                databaseName.Value = fd.SelectedItems(1)
                End If
        End With
            
        Set fd = Nothing
    
    End Sub
    
    ' Running the program/SQL query when the generate button(previously named run, but liked generate better) is clicked
    Private Sub runButton_Click()
        Dim conn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQL As String, databasePath As String
        Dim minAge As Integer, maxAge As Integer, i As Integer
        Dim positionfilter As String, teamfilter As String
        Dim ws As Worksheet
        Dim lastRow As Long, averageRow As Long, lastColumn As Long
        
        
        ' Inital Inputs
        databasePath = databaseName.Value
        minAge = CInt(minAgeTextBox.Value)
        maxAge = CInt(maxAgeTextBox.Value)
        
        ' ==== Error Handling / Input Validation ======================================+
        
        ' Making sure the age textboxes have numeric values
        If Not IsNumeric(minAgeTextBox.Value) Or Not IsNumeric(maxAgeTextBox.Value) Then
            MsgBox "Please enter valid numeric values for age range.", vbExclamation
            Exit Sub
        End If
        
        ' Making sure the min age is less than the max age
        If minAge > maxAge Then
            MsgBox "Minimum age cannot be greater than maximum age!", vbExclamation
            Exit Sub
        End If
        
        ' Making sure the minimum age is within the bounds (18 is the youngest player, 40 is the oldest)
        If minAge < 18 Or maxAge > 40 Then
            MsgBox "Age range must be between 18 and 40", vbExclamation
            Exit Sub
        End If
        
        ' If no database path is set, tell user to input the database
        If databasePath = "" Then
            MsgBox "Please select the database file!", vbExclamation
            Exit Sub
        End If
        
        ' =============================================================================+
        
        ' Function calls for dynamic SQL Query, gets whatever the user highlighted in the UserForm and appends it to the query
        positionfilter = positionFilterGetter()
        teamfilter = teamFilterGetter()
    
        ' Extremely long SQL Query, 90% of it is just stat feilds
        SQL = "SELECT Players.Player, Players.Pos, Players.Age, Players.Tm, " & _
              "Statistics.G, Statistics.GS, Statistics.MP, Statistics.ORB, Statistics.DRB, " & _
              "Statistics.TRB, Statistics.AST, Statistics.STL, Statistics.TOV, Statistics.PF, Statistics.PTS, " & _
              "Shooting.FG, Shooting.FGA, Shooting.[FG%], Shooting.[3P], Shooting.[3PA], Shooting.[3P%], " & _
              "Shooting.[2P], Shooting.[2PA], Shooting.[2P%], Shooting.[eFG%], Shooting.FT, Shooting.FTA, Shooting.[FT%] " & _
              "FROM (Players " & _
              "INNER JOIN Statistics ON Players.[PlayerID] = Statistics.[PlayerID]) " & _
              "INNER JOIN Shooting ON Players.[PlayerID] = Shooting.[PlayerID] " & _
              "WHERE Players.Age BETWEEN " & minAge & " AND " & maxAge & " " & positionfilter & " " & teamfilter
        
        ' Query wasn't working, ended up being because of a missing [] around a feild
        Debug.Print "SQL Query: " & SQL
    
        ' If there has been anything wrong so far with the query or databse
        On Error GoTo databaseError
        
        ' =============================================================================+
        ' Open connection
        With conn
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath & ";"
            .Open
        End With
        
        ' Run the SQL query
        rs.Open SQL, conn
        ' =============================================================================+
        
        ' Set new sheet to populate with the queried data!
        Set ws = Worksheets.Add
        ws.Name = "Filter Summary"
        
        ' Iterating from the start of the recordset columns to the end
        For i = 1 To rs.Fields.Count
            ' populating the worksheet with the names of the feilds
            ws.Cells(1, i).Value = rs.Fields(i - 1).Name
        Next i
        
        ' Bolding and adding color to the feild names
        With ws.Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Font.Color = RGB(0, 0, 255)
        End With
            
        '  populating the data into the worksheet from the recordset
        ws.Range("A2").CopyFromRecordset rs
        
        ' Autofitting the columns
        ws.Columns.AutoFit
        
        
        ' Find last row/column after data is printed, average row will be two rows under the lastrow
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        averageRow = lastRow + 2
        
        ' Making the word "Averages:" Bold
        With ws.Cells(averageRow, 1)
            .Value = "Averages: "
            .Font.Bold = True
        End With
    
        ' Loop through columns 2 to last column, calculate averages if numeric, skip columns B (2) and D (4) as they are not numbers
        For i = 2 To lastColumn
            If i <> 2 And i <> 4 Then ' Skip columns B and D
                On Error Resume Next
                ' Passing by reference as formula wouldn't work without it
                ws.Cells(averageRow, i).Formula = "=AVERAGE(" & ws.Cells(2, i).Address() & ":" & ws.Cells(lastRow, i).Address() & ")"
                On Error GoTo 0
            End If
        Next i
        
        ' Clean Up
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
        Exit Sub
        
databaseError:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Database Error"
        On Error GoTo 0
        
    End Sub
    
    'The position filter, it takes in what was highlighted by the user when selecting the position feilds to be appended into the SQL query
    Function positionFilterGetter() As String
        Dim positionfilter As String, i As Integer
        
        ' Start with empty String
        positionfilter = ""
        
        ' From start to finish, check if it is selected and append the value to the string.
        For i = 0 To positionListBox.ListCount - 1
            If positionListBox.Selected(i) And positionListBox.List(i) <> "None" Then
                ' Adds a comma at the end for the next word, will be removed at the end when loop is finished.
                positionfilter = positionfilter & "'" & positionListBox.List(i) & "',"
            End If
        Next i
        
        If positionfilter <> "" Then
            ' The Left(__, ___) gets rid of the one comma that is left behind that would cause a syntax error
            positionfilter = " AND Players.Pos IN (" & Left(positionfilter, Len(positionfilter) - 1) & ")"
        End If
        
        positionFilterGetter = positionfilter
    End Function
    
    ' Same code as position filter, but for the team listbox
    ' made modular to save space and make the run button code less complicated
    Function teamFilterGetter() As String
        Dim teamfilter As String, i As Integer
        
        teamfilter = ""
        For i = 0 To teamListBox.ListCount - 1
            If teamListBox.Selected(i) And teamListBox.List(i) <> "None" Then
                teamfilter = teamfilter & "'" & teamListBox.List(i) & "',"
            End If
        Next i
        
        If teamfilter <> "" Then
            teamfilter = " AND Players.Tm IN (" & Left(teamfilter, Len(teamfilter) - 1) & ")"
        End If
        
        teamFilterGetter = teamfilter
    End Function
