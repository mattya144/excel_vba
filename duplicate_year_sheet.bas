Attribute VB_Name = "Module1"
Sub deplicate_sheet()
Dim month As String: month = "ŒŽ"
Dim baseSheet As String
For i = 1 To 12
    If i = 1 Then
        Worksheets("Sheet1").Copy After:=Worksheets("Sheet1")
        ActiveSheet.Name = i & month
    Else
        baseSheet = (i - 1) & month
        Worksheets("Sheet1").Copy After:=Worksheets(baseSheet)
        ActiveSheet.Name = i & month
    End If
Next
End Sub
