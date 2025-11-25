Sub ExtractValuesFromAllSheets()
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim selectedRange As Range
    Dim area As Range
    Dim areaIndex As Long
    Dim numAreas As Long
    Dim rowIndex As Long
    Dim cellIndex As Long
    Dim maxCells As Long
    Dim cellCount As Long
    Dim cellValue As Variant
    Dim areaAddresses() As String
    Dim areaCellCounts() As Long
    
    ' Get the currently selected range
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select a range first!", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection
    
    ' Count the number of areas (handles non-contiguous ranges)
    numAreas = selectedRange.Areas.Count
    
    ' Store area addresses and cell counts (merged cells count as one)
    ReDim areaAddresses(1 To numAreas)
    ReDim areaCellCounts(1 To numAreas)
    maxCells = 0
    
    ' Helper function to count unique cells (merged cells count as one)
    Dim cell As Range
    Dim mergeAreas As Object
    Dim mergeAreaAddr As String
    
    For areaIndex = 1 To numAreas
        areaAddresses(areaIndex) = selectedRange.Areas(areaIndex).Address(False, False)
        
        ' Count unique cells, treating merged cells as one
        Set mergeAreas = CreateObject("Scripting.Dictionary")
        For Each cell In selectedRange.Areas(areaIndex).Cells
            ' Get the merge area address (for merged cells, this will be the merged range)
            mergeAreaAddr = cell.MergeArea.Address(False, False)
            ' Only count each unique merge area once
            If Not mergeAreas.Exists(mergeAreaAddr) Then
                mergeAreas.Add mergeAreaAddr, True
            End If
        Next cell
        areaCellCounts(areaIndex) = mergeAreas.Count
        
        If areaCellCounts(areaIndex) > maxCells Then
            maxCells = areaCellCounts(areaIndex)
        End If
    Next areaIndex
    
    ' Check if "Extracted_Values" sheet exists, otherwise create it
    On Error Resume Next
    Set outputWs = ThisWorkbook.Sheets("Extracted_Values")
    On Error GoTo 0
    
    If outputWs Is Nothing Then
        ' Create new worksheet if it doesn't exist
        Set outputWs = ThisWorkbook.Sheets.Add
        outputWs.Name = "Extracted_Values"
    Else
        ' Clear existing content if sheet exists
        outputWs.Cells.Clear
    End If
    
    ' Set up headers
    outputWs.Cells(1, 1).Value = "SheetName"
    For areaIndex = 1 To numAreas
        outputWs.Cells(1, areaIndex + 1).Value = areaAddresses(areaIndex)
    Next areaIndex
    
    ' Start writing data from row 2
    rowIndex = 2
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        ' Skip the output sheet itself
        If ws.Name <> "Extracted_Values" Then
            ' Collect unique cells for each area (merged cells count as one)
            Dim areaUniqueCells() As Object ' Array of collections/dictionaries
            ReDim areaUniqueCells(1 To numAreas)
            Dim uniqueCellIndex As Long
            
            ' First, collect all unique cells for each area
            For areaIndex = 1 To numAreas
                On Error Resume Next
                Set area = ws.Range(areaAddresses(areaIndex))
                On Error GoTo 0
                
                ' Create a collection to store unique cells for this area
                Set areaUniqueCells(areaIndex) = CreateObject("Scripting.Dictionary")
                
                If Not area Is Nothing Then
                    ' Collect unique cells (merged cells count as one)
                    Set mergeAreas = CreateObject("Scripting.Dictionary")
                    uniqueCellIndex = 0
                    
                    For Each cell In area.Cells
                        ' Get the merge area address
                        mergeAreaAddr = cell.MergeArea.Address(False, False)
                        ' Only add if we haven't seen this merge area before
                        If Not mergeAreas.Exists(mergeAreaAddr) Then
                            mergeAreas.Add mergeAreaAddr, True
                            uniqueCellIndex = uniqueCellIndex + 1
                            ' Store the top-left cell of the merge area
                            areaUniqueCells(areaIndex).Add uniqueCellIndex, cell.MergeArea.Cells(1)
                        End If
                    Next cell
                End If
            Next areaIndex
            
            ' Now write values row by row (one row per cell position)
            For cellIndex = 1 To maxCells
                ' Write sheet name in first column
                outputWs.Cells(rowIndex, 1).Value = ws.Name
                
                ' Write values from each area at this position
                For areaIndex = 1 To numAreas
                    Dim areaDict As Object
                    Set areaDict = areaUniqueCells(areaIndex)
                    
                    ' Check if this area has a cell at this position
                    If cellIndex <= areaCellCounts(areaIndex) And areaDict.Exists(cellIndex) Then
                        ' Get the cell value from the unique cell
                        Set cell = areaDict(cellIndex)
                        cellValue = cell.Value
                        If IsEmpty(cellValue) Then
                            cellValue = ""
                        End If
                        outputWs.Cells(rowIndex, areaIndex + 1).Value = cellValue
                    Else
                        ' No cell at this position for this area
                        outputWs.Cells(rowIndex, areaIndex + 1).Value = ""
                    End If
                Next areaIndex
                
                ' Move to next row
                rowIndex = rowIndex + 1
            Next cellIndex
        End If
    Next ws
    
    ' Format as table (optional - makes it look nicer)
    With outputWs
        .Columns.AutoFit
        With .Range(.Cells(1, 1), .Cells(1, numAreas + 1))
            .Font.Bold = True
            .Interior.Color = RGB(217, 225, 242) ' Light blue header
        End With
    End With
    
    ' Activate the output sheet
    outputWs.Activate
    
    MsgBox "Extracted values from " & numAreas & " range(s) across all sheets.", vbInformation
End Sub

