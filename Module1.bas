Attribute VB_Name = "Module1"
Sub A_Import_all()
    On Error GoTo Endit
    
    Dim fnd, sname, ftype As String
    
    Dim fs, f, f1, fc, s
    
    
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    


    'Displays a dialog box wherein the user enters the path to import text files
    fnd = InputBox("Specify the path", "Import text files")
    
    'Displays a dialog box wherein the user enters the file type extension
    ftype = InputBox("Select file type extension", "File extension")
  
    'End Macro if Cancel Button is Clicked or no Text is Entered
      If fnd = vbNullString Then Exit Sub
      
      
    folderspec = fnd
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    i = 0
    For Each f1 In fc
        If Right(f1.Name, 3) = ftype Then
        inFileName = f1.ParentFolder.Path & "\" & f1.Name
        i = i + 1
   
        Sheets.Add after:=ActiveSheet
        
        sname = inFileName
        
        asname = f1.Name
        asFname = Replace(asname, ftype, "")
        ActiveSheet.Name = asFname

        With ActiveSheet.QueryTables _
            .Add(Connection:="TEXT;" & sname, Destination:=ActiveCell)
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
            
        End With
 
        
        
    End If
    Next
    
    Sheets("Sheet1").Activate
    
    
    
Endit:
    If i = 0 Then
    'Displays a dialog box with the path and number of text files imported
    ret = MsgBox("No s2p file Found", vbOKOnly, "Bad Input")
    Else
    ret = MsgBox(i & " files found at " & inFileName, vbOKOnly)
    End If
    
End Sub
