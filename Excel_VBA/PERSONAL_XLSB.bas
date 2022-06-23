Option Explicit

Sub Unhide_All()
    Dim sh As Worksheet
    
    For Each sh In Worksheets
        sh.Visible = True
    Next sh

End Sub

' Leila's solution
Sub Auto_Table_Contents()
'this procedure adds a table of contents to your active sheet.
'It creates a hyperlink to each sheet and takes the header from cell A1 of every sheet

  Dim StartCell As Range               ' For InputBox to Select Range
  Dim EndCell As Range                 ' For Message Box as Info
  Dim Sh As Worksheet
  Dim ShName As String
  Dim MsgConfirm As VBA.VbMsgBoxResult ' For message box to confirm
     
  On Error Resume Next
  
  ' Prompt the user to select a cell location for TOC
  Set StartCell = Excel.Application.InputBox("Where do you want to insert the Table of Contents?" _
  & vbNewLine & "Please select a cell:", "Insert Table of Contents", , , , , , 8)
  If Err.Number = 424 Then Exit Sub                        ' If user selects cancel, exit sub
  On Error GoTo Handle
  
  Set StartCell = StartCell.Cells(1, 1)                    ' Start Cell for TOC - Force to single cell-reference if user input is range
  Set EndCell = StartCell.Offset(Worksheets.Count - 2, 1)  ' End Cell for ToC
  
  ' Get Overwritten Confirmation
  MsgConfirm = VBA.MsgBox("The values in cells:" & vbNewLine & StartCell.Address & " to " & EndCell.Address & _
  " could be overwritten. Would you like to continue?", vbOKCancel + vbDefaultButton2, "Confirmation Required!")
  If MsgConfirm = vbCancel Then Exit Sub
  
  ' Begin Populating ToC
  For Each Sh In Worksheets                                  ' Loop through all worksheets
    ShName = Sh.Name                                         ' Capture SheetName
    If ActiveSheet.Name <> ShName Then                       ' Include only if not activesheet
      If Sh.Visible = xlSheetVisible Then                    ' Include only if visible
        ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:= _
            "'" & ShName & "'!A1", TextToDisplay:=ShName     ' In ToC 1st col, add hyperlink with SheetName
            
        StartCell.Offset(0, 1).Value = Sh.Range("A1").Value  ' In ToC 2nd col, sheet title located in each sheet's A1
        Set StartCell = StartCell.Offset(1, 0)               ' Go to next row in TOC
      End If 'Sheet is visible
    End If   'Sheet is not activesheet
  Next Sh
  Exit Sub
  
Handle:
MsgBox "Unfortuntely, an error has occured"
End Sub
