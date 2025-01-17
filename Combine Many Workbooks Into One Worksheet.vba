Sub CombineManyWorkbooksIntoOneWorksheet()
    
    Dim strDirContainingFiles As String, strFile As String, _
        strFilePath As String
    Dim wbkDst As Workbook, wbkSrc As Workbook
    Dim wksDst As Worksheet, wksSrc As Worksheet
    Dim lngIdx As Long, lngSrcLastRow As Long, _
        lngSrcLastCol As Long, lngDstLastRow As Long, _
        lngDstLastCol As Long, lngDstFirstFileRow As Long
    Dim rngSrc As Range, rngDst As Range, rngFile As Range
    Dim colFileNames As Collection
    Set colFileNames = New Collection
    
    'Set references up-front
    strDirContainingFiles = "C:\Users\f4i65kw\Desktop\All_Sheets\" '<~ your folder
    'Set wbkDst = Workbooks.Add '<~ Dst is short for destination
    Set wbkDst = Workbooks(1)
    Set wksDst = wbkDst.ActiveSheet
    
    'Store all of the file names in a collection
    strFile = Dir(strDirContainingFiles & "\*.xlsx")
    Do While Len(strFile) > 0
        colFileNames.Add Item:=strFile
        strFile = Dir
    Loop
    
    ''CHECKPOINT: make sure colFileNames has the file names
    'Dim varDebug As Variant
    'For Each varDebug In colFileNames
    '    Debug.Print varDebug
    'Next varDebug
    
    'Now we can start looping through the "source" files
    'and copy their data to our destination sheet
    For lngIdx = 1 To colFileNames.Count
        
        'Assign the file path
        strFilePath = strDirContainingFiles & "\" & colFileNames(lngIdx)
        
        'Open the workbook and store a reference to the data sheet
        Set wbkSrc = Workbooks.Open(strFilePath)
        Set wksSrc = wbkSrc.Worksheets("Sheet2") '<~ change based on your Sheet name
        
        'Identify the last row and last column, then
        'use that info to identify the full data range
        lngSrcLastRow = LastOccupiedRowNum(wksSrc)
        lngSrcLastCol = LastOccupiedColNum(wksSrc)
        With wksSrc
            Set rngSrc = .Range(.Cells(1, 1), .Cells(lngSrcLastRow, _
                                                     lngSrcLastCol))
        End With
        
        ''CHECKPOINT: make sure we have the full source data range
        'wksSrc.Range("A1").Select
        'rngSrc.Select
        
        'If this is the first (1st) loop, we want to keep
        'the header row from the source data, but if not then
        'we want to remove it
        If lngIdx <> 1 Then
            Set rngSrc = rngSrc.Offset(1, 0).Resize(rngSrc.Rows.Count - 1)
        End If
        
        ''CHECKPOINT: make sure that we remove the header row
        ''from the source range on every loop that is not
        ''the first one
        'wksSrc.Range("A1").Select
        'rngSrc.Select
        
        'Copy the source data to the destination sheet, aiming
        'for cell A1 on the first loop then one past the
        'last-occupied row in column A on each following loop
        If lngIdx = 1 Then
            lngDstLastRow = 1
            Set rngDst = wksDst.Cells(1, 1)
        Else
            lngDstLastRow = LastOccupiedRowNum(wksDst)
            Set rngDst = wksDst.Cells(lngDstLastRow + 1, 1)
        End If
        rngSrc.Copy Destination:=rngDst '<~ this is the copy / paste
        
        'Almost done! We want to add the source file info
        'for each of the data blocks to our destination
        
        'On the first loop, we need to add a "Source Filename" column
        If lngIdx = 1 Then
            lngDstLastCol = LastOccupiedColNum(wksDst)
            wksDst.Cells(1, lngDstLastCol + 1) = "Source Filename"
        End If
        
        'Identify the range that we need to write the source file
        'info to, then write the info
        With wksDst
        
            'The first row we need to write the file info to
            'is the same row where we did our initial paste to
            'the destination file
            lngDstFirstFileRow = lngDstLastRow + 1
            
            'Then, we need to find the NEW last row on the destination
            'sheet, which will be further down (since we pasted more
            'data in)
            lngDstLastRow = LastOccupiedRowNum(wksDst)
            lngDstLastCol = LastOccupiedColNum(wksDst)
            
            'With the info from above, we can create the range
            Set rngFile = .Range(.Cells(lngDstFirstFileRow, lngDstLastCol), _
                                 .Cells(lngDstLastRow, lngDstLastCol))
                                 
            ''CHECKPOINT: make sure we have correctly identified
            ''the range where our file names will go
            'wksDst.Range("A1").Select
            'rngFile.Select
                                 
            'Now that we have that range identified,
            'we write the file name
            rngFile.Value = wbkSrc.name
            
        End With
        
        'Close the source workbook and repeat
        wbkSrc.Close SaveChanges:=False
        
    Next lngIdx
    
    'Let the user know that the combination is done!
    MsgBox "Data combined!"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last row
'OUTPUT      : Long, the last occupied row
'SPECIAL CASE: if Sheet is empty, return 1
Public Function LastOccupiedRowNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Row
        End With
    Else
        lng = 1
    End If
    LastOccupiedRowNum = lng
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last column
'OUTPUT      : Long, the last occupied column
'SPECIAL CASE: if Sheet is empty, return 1
Public Function LastOccupiedColNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Column
        End With
    Else
        lng = 1
    End If
    LastOccupiedColNum = lng
End Function



