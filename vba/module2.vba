Sub RunIngredientsUpdating()
Dim FilePath As String
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1

FilePath = Sheets("Overview").Cells(2, 2).Value
If IsFile(FilePath) = True Then
    commandString = Sheets("Overview").Cells(2, 2).Value & " " & ActiveWorkbook.Name & " " & Sheets("Overview").Cells(3, 2).Value
    wsh.Run "cmd.exe /S /C " & commandString, windowStyle, waitOnReturn
Else
    MsgBox ("Please download the Python Program before you run the Updating")
End If
End Sub


Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

