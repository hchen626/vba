# Key Takeways: Referencing Ranges

1. Different Methods to reference cells
2. Different methods to find the last row
  ```vba
  Range("K6").Value = Cells(Rows.Count,1).End(xlUp).Row
  Range("K10").Value = Range("A4").CurrentRegion.Rows.Count
  Range("K11").Value = Cells.SpecialCells(xlCellTypeLastCell).Row    'WARNING: If there's data below your table, that would be last row
  Range("K12").Value = Application.ActiveSheet.UsedRange.Rows.Count
  ```
3. Copy a variably sized range with the CurrentRegion Property
