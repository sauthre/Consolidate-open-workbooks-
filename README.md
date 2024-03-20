# Consolidate-open-workbooks-
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
