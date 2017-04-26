Private Sub Worksheet_Change(ByVal Target As Range)
Dim rngDV As Range
Dim oldVal As String
Dim newVal As String
If Target.Count > 1 Then GoTo exitHandler

On Error Resume Next
Set rngDV = Cells.SpecialCells(xlCellTypeAllValidation)
On Error GoTo exitHandler

If rngDV Is Nothing Then GoTo exitHandler

If Not Intersect(Target, rngDV) Is Nothing Then
	Application.EnableEvents = False
 	newVal = Target.Value
 	Application.Undo
 	oldVal = Target.Value
 	Target.Value = newVal
 	If Target.Column = 1 Then
		If oldVal <> "" And newVal <> ""  Then 
      		Target.Value = oldVal & ", " & newVal
      	End If
    End If
End If

exitHandler:
  Application.EnableEvents = True
End Sub

