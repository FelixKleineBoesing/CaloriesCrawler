Private Sub Worksheet_Change(ByVal Target As Range)

If Not Intersect(Target, Range("D2:D10000")) Is Nothing Then
    Call UpdateInformations
End If

End Sub
