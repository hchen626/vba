Sub Range_Lastrow()

Range("K6:K12").ClearContents
Range("A5", "A" & Cells(Rows.Count, 1).End(xlUp).Row).EntireRow.Interior.Color = Excel.Constants.xlNone

'Range("K6").Value = Range("A4").End(xlDown).Row                   'Find Last Row Here - Method 1
Range("K6").Value = Range("A" & Rows.Count).End(xlUp).Row          'Find Last Row Here - Method 2

Range("K7").Value = Range("A" & Rows.Count).End(xlUp).Row + 1      'Next Empty row #

'Range("K8").Value = Range("A4").End(xlToRight).Column             'Last column # - method 1
Range("K8").Value = Cells(4, Columns.Count).End(xlToLeft).Column   'Last column # - method 2

Range("K9").Value = Range("B10").CurrentRegion.Address             ' Current Region Address
Range("K10").Value = Range("B10").CurrentRegion.Rows.Count         ' # of Rows in the data set

Range("K11").Value = Cells.SpecialCells(xlCellTypeLastCell).Row    ' Last Row used - method 1
Range("K12").Value = ActiveSheet.UsedRange.Rows.Count              ' Last Row used - method 2

'Range("A" & Rows.Count).End(xlUp).EntireRow.Interior.Color = VBA.ColorConstants.vbRed

End Sub
