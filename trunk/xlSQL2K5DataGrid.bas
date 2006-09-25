Attribute VB_Name = "Module1"
'----------------------------------------------------------
'This program needs the following subroutines in the workbook:
' Private Sub Workbook_Open()
'    init
' End Sub
'And the following subroutines in the worksheet housing the table
'Private Sub Worksheet_Change(ByVal Target As Range)
'    Worksheet_ChangeSub Target
'End Sub
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    Worksheet_SelectionChangeSub Target
'End Sub
'---------------------------------------------------------
'
'---------------------------------------------------------
'This program needs a query table connected to a SQL 2005 database
' called testTags, which has a table called tags.
' The subroutine getTags() creates this query table
' Remember to change the WSID in the cnnStr constant,
' which should be the name of the machine where SQL 2005 resides.
'---------------------------------------------------------
'Tell the system to require declariation of variables.
Option Explicit
'String Variable used to store the state of the row
Public stateStr As String
'String Variable used to store the location of the cursor
Public locationStr As String
'Variable to store the last row number
Public rowNo As Integer
Public tblRows As Integer
Public tblColumns As Integer
'Variable to store the modify state of the row
Public rowModify As Boolean
'Variable to store the insert state of the row
Public rowNew As Boolean
'ODBC Conntection string for Query table
Public Const firstRow = 1
Public tblRng As Range
Public insertRng As Range
'Store the value of the tag field of current row
Public rowValue As String
'ODBC connectionstring for SQL Native Client used in the query table
Public Const cnnStr = "ODBC;DRIVER=SQL Native Client;" & _
    "SERVER=(local);UID=;Trusted_Connection=Yes;" & _
    "APP=Microsoft Office 2003;WSID=YOUR_SERVERNAME;DATABASE=testTags;"
'ADO OLEDB connectionstring for SQL Natvie Client
Public Const Connectionstring = "Provider=SQLNCLI;" & _
    "server=(local);Database=testtags;" & _
    "Integrated Security=SSPI;DataTypeCompatibility=80;"
Dim cnn As ADODB.Connection
'Initialize the program
Sub init()
    setModuleVariables
    setup_onKey
    Set cnn = New ADODB.Connection
End Sub
Sub beforeCloseWorkbook()
    cancel_onKey
    Set cnn = Nothing
End Sub
'Set up of some short-keys to run procedures
Sub setup_onKey()
    Application.OnKey "^{DEL}", "deleteTag"
    Application.OnKey "^{INSERT}", "insertProg"
End Sub
'Inserts a row and shifts rows down
Sub insertProg()
    Dim rad As Integer
    Dim kol As Integer
    Dim r As Range
    Set r = Range(ActiveCell.Address)
    rad = r.Row
    kol = r.Column
    Range(Cells(rad, 1), Cells(rad, 2)).Select
    Selection.Insert Shift:=xlDown
    Cells(rad, kol).Select
    rowNew = True
End Sub
Sub setModuleVariables()
    Set tblRng = Range("A1").CurrentRegion
    tblRows = tblRng.Rows.Count
    tblColumns = tblRng.Columns.Count
    Set insertRng = Range(Cells(tblRows + 1, 1).Address, _
    Cells(tblRows + 1, tblColumns).Address)
    rowNew = False
    rowModify = False
    stateStr = "Normal"
    rowNo = ActiveCell.Row
    rowValue = Cells(rowNo, 1).Value
    Application.EnableEvents = True
    Exit Sub
ERRORHANDLING:
    init
    Resume
End Sub
Sub Worksheet_ChangeSub(ByVal Target As Range)
    On Error GoTo ERRORHANDLING
    If stateStr = Empty Then
        stateStr = "Normal"
    End If
    If Not rowNew Then
    End If
    ' Exit if the target is the whole table (i.e the query table)
    If Target.Address = tblRng.Address Then
        Exit Sub
    End If
    ' If not in the grid
    If Intersect(Target, tblRng) Is Nothing Then
        If Intersect(Target, insertRng) Is Nothing Then
        Else
            rowNew = True
        End If
    Else
        'If the cursor is not in the title row
        If ActiveCell.Row > 1 Then
            'If a row isn't deleted then don't change the state
            If tblRng.Rows.Count < tblRows Then
            Else
                rowModify = True
            End If
        End If
    End If
    If rowNew Then
        stateStr = "insert"
    End If
    If rowModify Then
        If stateStr = "Normal" Then
            stateStr = "modify"
        Else
            stateStr = stateStr & " and modify"
        End If
    End If
    Application.StatusBar = locationStr + ", " + stateStr
End Sub
Sub Worksheet_SelectionChangeSub(ByVal Target As Range)
    Dim qt As QueryTable
    Set qt = ActiveSheet.QueryTables(1)
    On Error GoTo ERRORHANDLING
    locationStr = "tblRng"
    If stateStr = Empty Then
        stateStr = "Normal"
    End If
    If rowNo <> ActiveCell.Row Then
        Dim descriptionValue As String
        descriptionValue = Cells(rowNo, 2).Value
        If rowNew Then
            Dim tagValue As String
            tagValue = Cells(rowNo, 1).Value
            Application.EnableEvents = False
            'INSERT INTO tags
            insertTags tagValue, descriptionValue
            'Refresh Query table
            qt.Refresh
            rowNew = False
            'Reset the ranges
            setModuleVariables
        End If
        If rowModify Then
            Application.EnableEvents = False
            'UPDATE tags SET...
            'Bare oppdater andre felter enn indeks
            ' Get old value of tag
            updateTags rowValue, descriptionValue
            'Refresh Query table
            qt.Refresh
            'Reset ranges
            setModuleVariables
        End If
        stateStr = "Normal"
        rowNo = ActiveCell.Row
        'to be used if you change the index value
        rowValue = Cells(ActiveCell.Row, 1).Value
    End If
    If Intersect(Target, tblRng) Is Nothing Then
        ' And not int the insert range either
        If Intersect(insertRng, Target) Is Nothing Then
            locationStr = "outside tblRng and insertRng"
        Else
            locationStr = "insertRng"
        End If
    Else
        If ActiveCell.Row = 1 Then
            locationStr = "titleRng"
        Else
            locationStr = "tblRng"
        End If
    End If
    Application.StatusBar = locationStr + ", " + stateStr
    Exit Sub
ERRORHANDLING:
    init
    Resume
End Sub
'Deletes a tag
Sub deleteTag()
    Dim qt As QueryTable
    Set qt = ActiveSheet.QueryTables(1)
    Dim cmdCommand As ADODB.Command
    cnn.Open Connectionstring
    Dim sqlstring As String
    sqlstring = "delete from tags where tag = '" & _
                rowValue & "'"
    Set cmdCommand = New ADODB.Command
    Set cmdCommand.ActiveConnection = cnn
    With cmdCommand
        .CommandText = sqlstring
        .CommandType = adCmdText
        .Execute
    End With
    Set cmdCommand = Nothing
    cnn.Close
    qt.Refresh
    setModuleVariables
End Sub
'Updates the second column in a table
Sub updateTags(tagValue As String, descriptionValue As String)
    Dim cmdCommand As ADODB.Command
    cnn.Open Connectionstring
    Dim sqlstring As String
    sqlstring = "update tags set " & _
                Cells(firstRow, 2).Text & " = '" & _
                descriptionValue & _
                "' where tag = '" & rowValue & "'"
    Set cmdCommand = New ADODB.Command
    Set cmdCommand.ActiveConnection = cnn
    With cmdCommand
        .CommandText = sqlstring
        .CommandType = adCmdText
        .Execute
    End With
    Set cmdCommand = Nothing
    cnn.Close
End Sub
Sub insertTags(tagValue As String, descriptionValue As String)
    Dim cmdCommand As ADODB.Command
    cnn.Open Connectionstring
    'Set cnn = New ADODB.Connection
    Dim sqlstring As String
    sqlstring = "insert tags ( tag, description) " & _
                " values ( '" & tagValue & _
                "', '" & descriptionValue & "')"
    Set cmdCommand = New ADODB.Command
    Set cmdCommand.ActiveConnection = cnn
    With cmdCommand
        .CommandText = sqlstring
        .CommandType = adCmdText
        .Execute
    End With
    Set cmdCommand = Nothing
    cnn.Close
End Sub
Sub getTags()
    Dim sqlstring As String
    sqlstring = "Select * from tags"
    With ActiveSheet.QueryTables.Add(Connection:=cnnStr, _
        Destination:=Range("A1"), Sql:=sqlstring)
        .Refresh
    End With
End Sub
Sub cancel_onKey()
    Application.OnKey "^{DEL}"
    Application.OnKey "^{INSERT}"
End Sub
