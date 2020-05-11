Attribute VB_Name = "ApiCall"
Sub MainRun(URL As String)

Country = ActiveCell.Value

Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
URL = URL

'form request and send
objHttp.Open "GET", URL, False
objHttp.Send

''Response Object
If objHttp.Status = 200 Then
    Debug.Print "Success"
    Ut = objHttp.ResponseText
    Debug.Print Ut
    
Else
    Debug.Print "No Dice"
    Exit Sub
End If
    
'' Parse Json with Json Converter Library - https://github.com/VBA-tools/VBA-JSON
Dim Json As Object

Set Json = JsonConverter.ParseJson(Ut)

If URL = "https://api.covid19api.com/stats" Then
    Worksheets.Add
    Counter = 1
    For Each Item In Json
        Key = Item
        Cells(Counter, 1).Value = Item
        Cells(Counter, 2).Value = Json(Key)
        Counter = Counter + 1
    Next
    
ElseIf URL = "https://api.covid19api.com/summary" Then

    Worksheets.Add
    ColCount = 1
    RowCount = 2
    For I = 1 To Json("Countries").Count
        For Each Item In Json("Countries")(I)
            Key = Item
            Cells(1, ColCount).Value = Item
            Cells(RowCount, ColCount).Value = Json("Countries")(I)(Key)
            ColCount = ColCount + 1
        Next Item
    RowCount = RowCount + 1
    ColCount = 1
    Next I

ElseIf URL = "https://api.covid19api.com/dayone/country/" & Country Then

    Worksheets.Add
    ColCount = 1
    RowCount = 2
    
    For Each Item In Json
        For Each it In Item
        Key = it
        Cells(1, ColCount).Value = it
        Cells(RowCount, ColCount).Value = Json(RowCount - 1)(Key)
        ColCount = ColCount + 1
        Next it
    RowCount = RowCount + 1
    ColCount = 1
    Next Item


    
End If

End Sub

Sub Stats(Ribbon As IRibbonControl)
URL = "https://api.covid19api.com/stats"
MainRun (URL)

End Sub

Sub Summary(Ribbon As IRibbonControl)
URL = "https://api.covid19api.com/summary"
MainRun (URL)

End Sub

Sub CountryStats(Ribbon As IRibbonControl)
Country = ActiveCell.Value

If ActiveCell.Column <> 3 Then
    MsgBox "You must select Country from Column 3 to run detail"
    Exit Sub
End If

URL = "https://api.covid19api.com/dayone/country/" & Country
MainRun (URL)
End Sub
