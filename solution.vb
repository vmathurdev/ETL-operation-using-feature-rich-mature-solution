' Update Function

Sub dataconfigupdate()

Dim cnn1 As New ADODB.Connection
    Dim mrs As New ADODB.Recordset
    Dim iCols As Integer
    Const DRIVER = "{SQL Server}"
    Dim sserver As String
    Dim ddatabase As String
    Dim iRN As Integer
    
    
    
    Worksheets("Settings").Range("H32") = Now
    
    sserver = Worksheets("Settings").Range("D11").Value
    ddatabase = Worksheets("Settings").Range("D12").Value

    Set cnn1 = New ADODB.Connection
    cnn1.ConnectionString = "driver= " & DRIVER & ";server=" & sserver & ";database=" & ddatabase & ""
      cnn1.ConnectionTimeout = 30
      cnn1.Open
      With Worksheets("EnterConfig")
            
        
      cnn1.Execute "delete from [dbo].[ssisPckgVariables]"
      
'Skip the header row
        iRN = 2
            
        'Loop until empty cell
        Do Until .Cells(iRN, 1) = ""
            sPackageID = .Cells(iRN, 1)
            sVariableName = .Cells(iRN, 2)
            sVariableValue = .Cells(iRN, 3)
            sSegment = .Cells(iRN, 4)
            sActive = .Cells(iRN, 5)
            sSegmentFlow = .Cells(iRN, 6)
            sTargetTable = .Cells(iRN, 7)
            sDescription = .Cells(iRN, 8)
                
             
                
                
            'Generate and execute sql statement to import the excel rows to SQL Server table
            cnn1.Execute "insert into [dbo].[ssisPckgVariables] (PackageID,VariableName,VariableValue,Segment,Active,SegmentFlow,TargetTable,Description) values ('" & sPackageID & "', '" & sVariableName & "', '" & sVariableValue & "', '" & sSegment & "', '" & sActive & "', '" & sSegmentFlow & "',  '" & sTargetTable & "', '" & sDescription & "')"
 
            iRN = iRN + 1
        Loop
     
       cnn1.Close
       End With
       
       MsgBox "Data updated."
       
             
       Worksheets("Settings").Range("D32") = Now
    

End Sub

' Sub to fetch the Aggregated (insert/update) data
Sub SavedConfiguration()

    Dim cnn1 As New ADODB.Connection
    Dim mrs As New ADODB.Recordset
    Dim iCols As Integer
    Const DRIVER = "{SQL Server}"
    Dim sserver As String
    Dim ddatabase As String
    
    Worksheets("Settings").Range("H35") = Now
    
    sserver = Worksheets("Settings").Range("D11").Value
    ddatabase = Worksheets("Settings").Range("D12").Value

    Set cnn1 = New ADODB.Connection
    cnn1.ConnectionString = "driver= " & DRIVER & ";server=" & sserver & ";database=" & ddatabase & ""
      cnn1.ConnectionTimeout = 30
      cnn1.Open

    sQry = Worksheets("SQL-COMMON").Range("B3").Value

    mrs.Open sQry, cnn1

    

    For iCols = 0 To mrs.Fields.Count - 1
        Worksheets("SavedConfig").Cells(1, iCols + 1).Value = mrs.Fields(iCols).Name
    Next


    Worksheets("SavedConfig").Range("A2").CopyFromRecordset mrs
    
    

    mrs.Close
    cnn1.Close

    Worksheets("Settings").Range("D35") = Date + Time

End Sub

'Sub to insert new records in the table
Sub insertConfiguration()

    Dim cnn1 As New ADODB.Connection
    Dim mrs As New ADODB.Recordset
    Dim iCols As Integer
    Const DRIVER = "{SQL Server}"
    Dim sserver As String
    Dim ddatabase As String
    Dim iRowNo As Integer
    
      If IsEmpty(Range("A1:Z1000").Value) = True Then
      MsgBox "Cell A1 is empty"
   End If

    
    Worksheets("Settings").Range("H32") = Now
    
    sserver = Worksheets("Settings").Range("D11").Value
    ddatabase = Worksheets("Settings").Range("D12").Value

    Set cnn1 = New ADODB.Connection
    cnn1.ConnectionString = "driver= " & DRIVER & ";server=" & sserver & ";database=" & ddatabase & ""
      cnn1.ConnectionTimeout = 30
      cnn1.Open
      
   
      With Worksheets("EnterConfig")
      
'Skip the header row
        iRowNo = 201
            
        'Loop until empty cell
        Do Until .Cells(iRowNo, 1) = ""
            sPackageID = .Cells(iRowNo, 1)
            sVariableName = .Cells(iRowNo, 2)
            sVariableValue = .Cells(iRowNo, 3)
            sSegment = .Cells(iRowNo, 4)
            sActive = .Cells(iRowNo, 5)
            sSegmentFlow = .Cells(iRowNo, 6)
            sTargetTable = .Cells(iRowNo, 7)
            sDescription = .Cells(iRowNo, 8)
                
            'Generate and execute sql statement to import the excel rows to SQL Server table
            cnn1.Execute "insert into [dbo].[ssisPckgVariables] (PackageID,VariableName,VariableValue,Segment,Active,SegmentFlow,TargetTable,Description) values ('" & sPackageID & "', '" & sVariableName & "', '" & sVariableValue & "', '" & sSegment & "', '" & sActive & "', '" & sSegmentFlow & "',  '" & sTargetTable & "', '" & sDescription & "')"
 
            iRowNo = iRowNo + 1
        Loop
     
       cnn1.Close
       End With
       
       MsgBox "Data imported."
       Worksheets("Settings").Range("D32") = Now
    
End Sub


' Sub to fetch the data from SQL server
Sub EnterConfiguration()

    Dim cnn1 As New ADODB.Connection
    Dim mrs As New ADODB.Recordset
    Dim iCols As Integer
    Const DRIVER = "{SQL Server}"
    Dim sserver As String
    Dim ddatabase As String
    
    Worksheets("Settings").Range("H30") = Now  'Start time
    
    sserver = Worksheets("Settings").Range("D11").Value     'server name
    ddatabase = Worksheets("Settings").Range("D12").Value   'database name


    'Establishing the connection
    Set cnn1 = New ADODB.Connection
    cnn1.ConnectionString = "driver= " & DRIVER & ";server=" & sserver & ";database=" & ddatabase & ""
      cnn1.ConnectionTimeout = 30
      cnn1.Open

    'Query to execute the fetch request
    sQry = Worksheets("SQL-COMMON").Range("B3").Value

    mrs.Open sQry, cnn1

    
    'Print the results with headers
    For iCols = 0 To mrs.Fields.Count - 1
        Worksheets("EnterConfig").Cells(1, iCols + 1).Value = mrs.Fields(iCols).Name
    Next

    ' Save the records
    Worksheets("EnterConfig").Range("A2").CopyFromRecordset mrs
    
    

    mrs.Close
    cnn1.Close
    
        
    Worksheets("Settings").Range("D30") = Date + Time  'Stop time
    

End Sub


