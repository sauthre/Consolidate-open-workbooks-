
Sub CopyRowsWithText()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim i As Long

    ' Set source and target sheets
    Set sourceSheet = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your source sheet name
    Set targetSheet = ThisWorkbook.Sheets("Sheet2") ' Change "Sheet2" to your target sheet name

    ' Find the last row with data in column A of source sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in column A and copy if there's text
    For i = 1 To lastRow
        If sourceSheet.Cells(i, 1).Value <> "" Then
            ' Find the next empty row in target sheet
            targetRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1
            ' Copy entire row from source to target sheet
            sourceSheet.Rows(i).Copy targetSheet.Rows(targetRow)
        End If
    Next i
End Sub




Consolidate-open-workbooks-
Sub CopyDataFromAllWorkbooks()
    Dim srcWorkbook As Workbook
    Dim destWorkbook As Workbook
    Dim sourceRange As Range
    Dim destRange As Range
    
    ' Set the destination range in the active sheet
    Set destWorkbook = ActiveWorkbook
    Set destRange = destWorkbook.Sheets(1).Range("A19:Q30")
    
    ' Loop through all open workbooks
    For Each srcWorkbook In Workbooks
        ' Exclude the destination workbook
        If srcWorkbook.Name <> destWorkbook.Name Then
            ' Set the source range in the first sheet of each open workbook
            Set sourceRange = srcWorkbook.Sheets(1).Range("A19:Q30")
            
            ' Copy data from source range to destination range
            destRange.Value = sourceRange.Value
            
            ' Move to the next empty row in the destination range
            Set destRange = destRange.Offset(destRange.Rows.Count, 0)
        End If
    Next srcWorkbook
End Sub
