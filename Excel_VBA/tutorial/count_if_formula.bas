Option Explicit

Sub Count_Formulas()
    Dim rng As Range
    Dim cel As Range
    Dim i As Long
    
    Set rng = Application.ActiveSheet.UsedRange
    'Debug.Print rng.Address
    
    i = 0
    For Each cel In rng.Cells
        If cel.HasFormula = True Then i = i + 1
    Next cel
        
    Range("B6").Value = i
    
End Sub

Sub Count_If_Formula()
' Leila Gharani - approach
' For the used range on this sheet, use a macro to get the number of cells that have formulas in them
' Place this value inside cell B6. Create a button and assign the macro to it.

    Dim cell As Range
    Dim CountFormula As Long
    
    
    For Each cell In ActiveSheet.UsedRange
        If cell.HasFormula Then
            CountFormula = CountFormula + 1
        End If
    Next cell
    Range("B6").Value = CountFormula

End Sub
