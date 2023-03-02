Sub UpdateInformations()
Dim ColumnStartingIngreds As Integer
Dim comment_index As Integer
Dim IngredientsArray() As Variant
Dim LastRow As Integer
Dim Name As String
Dim amount As Integer
Dim unit As String
Dim MaxNumberIngredients As Integer
Dim j As Integer
Dim LastColumnName As String
Dim IngredientsSheetName As String
Dim IngredientsColumnName As String


LastColumnName = "Comment"
IngredientsSheetName = "Ingredients"
IngredientsColumnName = "Ingredients"

On Error Resume Next

Application.Calculation = xlManual

Set NotCorrectRows = CreateObject("System.Collections.ArrayList")
Set NewIngredientsArray = CreateObject("System.Collections.ArrayList")

LastRow = Worksheets(IngredientsSheetName).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
RangeZutaten = "A1:A" & LastRow
IngredientsArray = Worksheets(IngredientsSheetName).Range("A1:A" & LastRow).Value2
If Cells(1, 4) <> IngredientsColumnName Then
    MsgBox "Exit Macro since the Layout is changed. Fourth Column was Zutaten"
    Exit Sub
End If

For j = 1 To 20
    If Cells(1, j).Value = LastColumnName Then
        ColumnStartingIngredsFirst = j + 1
    End If
Next j



i = 2
MaxNumberIngredients = 0
LastNumberError = 0
Do While Cells(i, 2).Value <> ""
    ColumnStartingIngreds = ColumnStartingIngredsFirst
    ingreds = Cells(i, 4).Value
    ingreds_array = Split(ingreds, vbLf)
    j = 0
    For Each Item In ingreds_array
        If Left(Item, 1) = "-" Then
            splitted_ingred = Split(Item, " ")
            amount_unit = splitted_ingred(0)
            comment_index = FindCommentIndex(splitted_ingred)
            Name = Join(SliceCopy(splitted_ingred, 1, comment_index))
            amount = NumericString((amount_unit))
            unit = CharString((amount_unit))
            Cells(i, ColumnStartingIngreds).Value = Name
            Cells(i, ColumnStartingIngreds + 1).Value = amount
            Cells(i, ColumnStartingIngreds + 2).Value = unit
            Cells(i, ColumnStartingIngreds + 3).Formula = "=" & Col_Letter(ColumnStartingIngreds + 1) & i & "* VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,5,FALSE) / VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,10,FALSE)"
            Cells(i, ColumnStartingIngreds + 4).Formula = "=" & Col_Letter(ColumnStartingIngreds + 1) & i & "* VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,6,FALSE) / VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,10,FALSE)"
            Cells(i, ColumnStartingIngreds + 5).Formula = "=" & Col_Letter(ColumnStartingIngreds + 1) & i & "* VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,7,FALSE) / VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,10,FALSE)"
            Cells(i, ColumnStartingIngreds + 6).Formula = "=" & Col_Letter(ColumnStartingIngreds + 1) & i & "* VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,8,FALSE) / VLOOKUP(" & Col_Letter(ColumnStartingIngreds) & i & "," & IngredientsSheetName & "!$A:$J,10,FALSE)"
            ColumnStartingIngreds = ColumnStartingIngreds + 7
            If Not IsInArray(Name, IngredientsArray) And Not NewIngredientsArray.contains(Name) Then
                NewIngredientsArray.Add Name
            End If
            If Err.Number > LastNumberError Then
                NotCorrectRows.Add i
            End If
            LastNumberError = Err.Number
            j = j + 1
        End If
    Next Item
    MaxNumberIngredients = WorksheetFunction.Max(j, MaxNumberIngredients)
    i = i + 1
Loop

For Each NewIngredient In NewIngredientsArray
    Sheets(IngredientsSheetName).Cells(LastRow, 1).Value = NewIngredient
    LastRow = LastRow + 1
Next NewIngredient


Application.Calculation = xlAutomatic
MsgBox ("The following rows are not correct formatted:" & NotCorrectRows)

End Sub

Sub DownloadProgram()

Dim myURL As String
myURL = "https://github.com/FelixKleineBoesing/CaloriesCrawler/blob/master/dist/main.exe"

Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False
WinHttpReq.send

If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile Cells(1, 1).Value, 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
End If

End Sub



Public Function IsInArray(stringToBeFound As String, arr As Variant, Optional column As Integer = 1) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, column) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function


Function Col_Letter(lngCol As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function


Function FindCommentIndex(MyArray) As Integer
LengthArray = GetLengthOfArray(MyArray)
FindCommentIndex = LengthArray
For i = 0 To LengthArray - 1
    If Left(MyArray(i), 1) = "#" Then
        FindCommentIndex = i
    End If
Next i

End Function

Function NextEmptyCol() As Integer
    NextEmptyCol = 1
    i = 1
    Do Until NextEmptyCol > 1
        If Cells(1, i).Value = "" Then NextEmptyCol = i
        i = i + 1
    Loop

End Function


Function NextEmptyRow(SheetName As String, Optional PivotCol As Integer = 1) As Integer
    ws = ActiveWorkbook.Worksheets(SheetName)
    i = 1
    NextEmptyRow = 1
    Do Until NextEmptyRow > 1
        If ws.Cells(i, PivotCol).Value = "" Then NextEmptyRow = i
        i = i + 1
    Loop

End Function

Function NumericString(WholeString As String) As Integer
    Dim regex As New RegExp

    regex.Pattern = "([0-9]+)"
    Set Founds = regex.Execute(WholeString)
    NumericString = CInt(Founds(0).Value)

End Function

Function CharString(WholeString As String) As String
    Dim regex As New RegExp

    regex.Pattern = "([a-z]+)"
    Set Founds = regex.Execute(WholeString)
    CharString = Founds(0).Value

End Function

Function SliceCopy(MyArray, StartInt As Integer, Optional EndInt As Integer = 0)
LengthArray = GetLengthOfArray(MyArray)
LengthNewArray = LengthArray - StartInt - (LengthArray - EndInt)

Dim NewArray() As Variant
ReDim NewArray(LengthNewArray - 1)
If EndInt = 0 Then EndInt = LengthArray

For i = 0 To LengthNewArray - 1
    NewArray(i) = MyArray(StartInt + i)
Next i

SliceCopy = NewArray

End Function


Function GetLengthOfArray(MyArray) As Integer
GetLengthOfArray = UBound(MyArray, 1) - LBound(MyArray, 1) + 1
End Function

