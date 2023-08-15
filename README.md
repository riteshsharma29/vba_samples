# vba_samples

This repo contains VBA code snippets.

ActiveWorkbook.Worksheets(sheetno).UsedRange.Rows.Count
Sub REMOVE_BLANKS()
Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub
