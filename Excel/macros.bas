Sub process1QRCodeInput()
    saveData (getInput())
End Sub

Sub process6QRCodeInput()
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
End Sub

Public Function getInput()
    getInput = InputBox("Scan QR Code", "Match Scouting Input")
End Function

Sub testSaveData()
    saveData ("s=fff;e=1234;l=qm;m=1234;r=r1;t=1234;as=;ae=Y;al=2;ao=2;ai=1;aa=Y;at=N;ax=Y;lp=2;op=1;ip=3;rc=pass;f=0;pc=pass;ss=;c=pass;b=N;ca=x;cb=x;cs=slow;p=N;ds=x;dr=x;pl=x;tr=N;wd=N;if=N;d=N;to=N;be=N;cf=N")
End Sub

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr)
End Function

Sub saveData(inp As String)
    Dim fields
    Dim par
    Dim value
    Dim key
    Dim table As ListObject
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim mapper
    Set mapper = CreateObject("Scripting.Dictionary")
    Dim data
    Set data = CreateObject("Scripting.Dictionary")
    Dim tableName As String
    tableName = "ScoutingData"

    ' Set up map
    ' Fields for every year
    mapper.Add "s", "scouter"
    mapper.Add "e", "eventCode"
    mapper.Add "l", "matchLevel"
    mapper.Add "m", "matchNumber"
    mapper.Add "r", "robot"
    mapper.Add "t", "teamNumber"

    ' Additional custom mapping
    'mapper.Add "f", "fouls"
    'mapper.Add "c", "climb"
    'mapper.Add "dr", "defenseRating"
    'mapper.Add "d", "died"
    'mapper.Add "to", "tippedOver"
    'mapper.Add "cf", "cardFouls"
    'mapper.Add "co", "comments"
    
    If inp = "" Then
        Exit Sub
    End If

    'MsgBox (inp)
    
    fields = Split(inp, ";")
    If ArrayLen(fields) > 0 Then
        Dim i As Integer
        Dim str

        i = 0

        For Each str In fields
            par = Split(str, "=")
            key = par(0)
            value = par(1)
            If mapper.Exists(key) Then
                key = mapper(key)
            End If
            data.Add key, value
        Next

        tableexists = False
        
        Dim tbl As ListObject
        Dim sht As Worksheet

        'Loop through each sheet and table in the workbook
        For Each sht In ThisWorkbook.Worksheets
            For Each tbl In sht.ListObjects
                If tbl.Name = tableName Then
                    tableexists = True
                    Set table = tbl
                    Set ws = sht
                End If
            Next tbl
        Next sht

        If tableexists Then
            'Set table = ws.ListObjects(tableName)
        Else
            Dim tablerange As Range
            ws.ListObjects.Add(xlSrcRange, Range("A1:AO1"), , xlYes).Name = tableName
            i = 0
            Set table = ws.ListObjects(tableName)
            For Each key In data.Keys
                table.Range(i + 1) = key
                i = i + 1
            Next
        End If

        Dim newrow As ListRow
    
        Set newrow = table.ListRows.Add
        
        For Each str In data.Keys
            newrow.Range(table.ListColumns(str).Index) = data(str)
        Next
    End If
End Sub

Sub saveToCSV()

  ' Retrieve data from the "ScoutingData" table
  Dim table As ListObject
  Dim ws As Worksheet
  Set ws = ActiveSheet
  Set table = ws.ListObjects("ScoutingData") ' Assuming the table exists

  ' Create file system object
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Define filename and path (replace with your desired location)
  Dim filePath As String
  filePath = "C:\ScoutingData.csv" ' Adjust as needed

  ' Create the CSV file
  Dim file As Object
  Set file = fso.CreateTextFile(filePath, True)

  ' Write header row
  For i = 1 To table.ListColumns.Count
    file.Write table.HeaderRowRange(i) & ","
  Next i
  file.Write vbCrLf ' Add a new line

  ' Write data rows
  For Each Row In table.ListRows
    For i = 1 To table.ListColumns.Count
      file.Write Row.Range(i) & ","
    Next i
    file.Write vbCrLf ' Add a new line
  Next Row

  ' Close the file
  file.Close

  ' Optional message box confirmation
  MsgBox "Data exported successfully to " & filePath, vbInformation

End Sub

Sub ClearAllData()

  ' Get the active worksheet
  Dim ws As Worksheet
  Set ws = ActiveSheet
  
  ' Prompt for confirmation
  If MsgBox("Are you sure you want to clear all data from '" & ws.Name & "'?", vbYesNo) = vbNo Then
    Exit Sub ' Exit the macro if user clicks "No"
  End If
  
  ' Prompt for confirmation
  If MsgBox("Are you sure you want to clear all data from '" & ws.Name & "'?", vbYesNo) = vbNo Then
    Exit Sub ' Exit the macro if user clicks "No"
  End If
  
  ' Prompt for confirmation
  If MsgBox("Are you sure you want to clear all data from(Last time i promise!) '" & ws.Name & "'?", vbYesNo) = vbNo Then
    Exit Sub ' Exit the macro if user clicks "No"
  End If
  
  

  ' Clear all contents, formats, and comments
  ws.Cells.ClearContents
  'ws.Cells.ClearComments
  'ws.Cells.ClearFormats
  
End Sub