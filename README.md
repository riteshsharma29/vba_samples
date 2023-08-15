# vba_samples

This repo contains VBA code snippets.

ActiveWorkbook.Worksheets(sheetno).UsedRange.Rows.Count
Sub REMOVE_BLANKS()<br>
Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete<br>
End Sub<br>
