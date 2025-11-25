Sub ExtractValuesFromAllSheets()
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim cell As Range
    Dim selectedRange As Range
    Dim selectedRangeAddress As String
    Dim rng As Range
    Dim results As Collection
    Dim item As Variant
    Dim nextCol As Long
    Dim sheetName As String
    Dim nextRow As Long
    Dim cellValue As Variant

    ' Get the currently selected range
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select a range first!", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection
    selectedRangeAddress = selectedRange.Address

    ' Prepare collection to store results during the loop
    Set results = New Collection

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        Set rng = ws.Range(selectedRangeAddress)
        On Error GoTo 0

        If Not rng Is Nothing Then
            ' Iterate through each cell in the range on this sheet
            ' This handles merged cells properly - merged cells will only be counted once
            For Each cell In rng.Cells
                ' Get the value from the cell (works correctly for merged cells)
                cellValue = cell.Value
                
                ' Only capture if the cell is not empty
                If cellValue <> "" And Not IsEmpty(cellValue) Then
                    results.Add Array(ws.Name, cellValue)
                End If
            Next cell
        End If
        Set rng = Nothing
    Next ws

    ' Check if "Extracted Values" sheet exists, otherwise create it
    On Error Resume Next
    Set outputWs = ThisWorkbook.Sheets("Extracted Values")
    On Error GoTo 0
    
    If outputWs Is Nothing Then
        ' Create new worksheet if it doesn't exist
        Set outputWs = ThisWorkbook.Sheets.Add
        outputWs.Name = "Extracted Values"
        ' Set initial header
        outputWs.Cells(1, 1).Value = "Sheet Name"
    End If
    
    ' Find the next empty column
    nextCol = 1
    Do While outputWs.Cells(1, nextCol).Value <> ""
        nextCol = nextCol + 1
    Loop
    
    ' Set header for the new column (use the range address)
    outputWs.Cells(1, nextCol).Value = selectedRangeAddress
    
    ' Write collected results
    ' Within a single run: allow multiple rows per sheet name
    ' Across runs: match existing sheet names to existing rows (don't duplicate)
    Dim lastRow As Long
    Dim foundRow As Long
    Dim i As Long
    Dim usedRowsInThisRun As Object ' Dictionary to track which rows we've used for each sheet name in this run
    
    ' Create dictionary to track used rows (using Scripting.Dictionary)
    Set usedRowsInThisRun = CreateObject("Scripting.Dictionary")
    
    ' Get the last row before we start adding
    lastRow = outputWs.Cells(outputWs.Rows.Count, 1).End(xlUp).Row
    
    For Each item In results
        sheetName = item(0)
        foundRow = 0
        
        ' Check if we've already used a row for this sheet name in THIS run
        If usedRowsInThisRun.Exists(sheetName) Then
            ' We've already used a row for this sheet name in this run, so create a new row
            lastRow = outputWs.Cells(outputWs.Rows.Count, 1).End(xlUp).Row
            foundRow = lastRow + 1
            outputWs.Cells(foundRow, 1).Value = sheetName
            ' Track that we've used this row for this sheet name
            usedRowsInThisRun(sheetName) = usedRowsInThisRun(sheetName) & "," & foundRow
        Else
            ' First time seeing this sheet name in this run - check if it exists from previous runs
            For i = 1 To lastRow
                If outputWs.Cells(i, 1).Value = sheetName Then
                    ' Found existing row - check if it already has a value in the new column
                    If outputWs.Cells(i, nextCol).Value = "" Then
                        ' Use this existing row
                        foundRow = i
                        Exit For
                    End If
                End If
            Next i
            
            ' If no suitable existing row found, create a new one
            If foundRow = 0 Then
                lastRow = outputWs.Cells(outputWs.Rows.Count, 1).End(xlUp).Row
                foundRow = lastRow + 1
                outputWs.Cells(foundRow, 1).Value = sheetName
            End If
            
            ' Track that we've used this row for this sheet name
            usedRowsInThisRun(sheetName) = foundRow
        End If
        
        ' Add value to the appropriate row in the new column
        outputWs.Cells(foundRow, nextCol).Value = item(1)
    Next item

    MsgBox "Values extracted to '" & outputWs.Name & "' sheet.", vbInformation
End Sub

