1. Count_If_Formula
  - Concepts: ActiveSheet.UsedRange, Range

# Key Takeways

1. Count_If_Formula
-   - Concepts: ActiveSheet.UsedRange, Range
2. Faster
  ```vba
  'Suppress
  With Application
    .StatusBar = "Short wait..."        ' Status bar to let user know the macro is running
    .ScreenUpdating = False             ' Screen Flickering
    .DisplayAlerts = False              ' Suppress Excel alerts when delete wkshts e.g.
    .Calculation = xlCalculationManual  ' Formula Calculations
  End With
  
  'Restore
    With Application
      .ScreenUpdating = True
      .DisplayAlerts = True
      .StatusBar = ""
      .Calculation = xlCalculationAutomatic
      .CutCopyMode = False              ' In case you used paste special and turn off dotted box
  End With
  ```
3. Iterating example
```vba
Sub Replace_Formula()
  ' Update cell formulas to have iferror
  Dim cell As Range
  Dim FormulaRange As Range
  
  Set FormulaRange = Cells.SpecialCells(xlCellTypeFormulas)
  For Each cell In FormulaRange
    cell.Formula = "=iferror(" & VBA.Mid(cell.Formula, 2) & ","""")"
  Next cell

End Sub
```
