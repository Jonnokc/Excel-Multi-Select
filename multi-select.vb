Private Sub Worksheet_Change(ByVal Target As Range)
Dim rngDV As Range
Dim oldVal As String
Dim newVal As String
If Target.Count > 1 Then GoTo exitHandler

On Error Resume Next
Set rngDV = Cells.SpecialCells(xlCellTypeAllValidation)
On Error GoTo exitHandler

If rngDV Is Nothing Then GoTo exitHandler

If Intersect(Target, rngDV) Is Nothing Then
Else
  Application.EnableEvents = False
  newVal = Target.Value
  Application.Undo
  oldVal = Target.Value
  Target.Value = newVal
  If Target.Column = 1 Then
    If oldVal = "" Then
      Else
      If newVal = "" Then
      Else
      Target.Value = oldVal _
        & ", " & newVal
      End If
    End If
  End If
End If

exitHandler:
  Application.EnableEvents = True
End Sub

