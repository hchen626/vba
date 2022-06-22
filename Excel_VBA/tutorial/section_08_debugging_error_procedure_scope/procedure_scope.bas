Option Explicit

Private Sub Entry_Point()
  'Suppress
  With Application
      .StatusBar = "Short wait..."        ' Status bar to let user know the macro is running
      .ScreenUpdating = False             ' Screen Flickering
      .DisplayAlerts = False              ' Suppress Excel alerts when delete wkshts e.g.
      .Calculation = xlCalculationManual  ' Formula Calculations
  End With
End Sub

Private Sub Exit_Point()
  'Restore
  With Application
      .ScreenUpdating = True
      .DisplayAlerts = True
      .StatusBar = ""
      .Calculation = xlCalculationAutomatic ' Turn back to default mode automatic sheet calcs
      .CutCopyMode = False                  ' In case you used paste special and turn off dotted box
  End With
End Sub


Sub Do_Stuff()

  Dim ShNew As Worksheet
  Dim cell As Range
  
  Call Entry_Point ' Call entry sub procedure
  
  Set ShNew = Worksheets.Add
  For Each cell In ShNew.Range("A1:A100000")
      cell.Value = 10
  Next cell
  
  ShNew.Delete
  Sheet8.Select
    
  Exit_Point       ' Call exit sub producedure
  
End Sub
