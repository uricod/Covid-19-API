Attribute VB_Name = "ApiCall"
Public Json As Object

Sub MainRun(URL As String)
On Error Resume Next
Country = ActiveCell.Value
On Error GoTo 0
Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
URL = URL

'form request and send
objHttp.Open "GET", URL, False
objHttp.Send

''Response Object
If objHttp.Status = 200 Then
    Debug.Print "Success"
    ut = objHttp.ResponseText
    Debug.Print ut
    
Else
    Debug.Print "No Dice"
    Exit Sub
End If
    
'' Parse Json with Json Converter Library - https://github.com/VBA-tools/VBA-JSON
Set Json = JsonConverter.ParseJson(ut)

End Sub

Sub Stats(Ribbon As IRibbonControl)
URL = "https://api.covid19api.com/stats"
MainRun (URL)

'' Handle Response
Worksheets.Add
Counter = 1
For Each Item In Json
    Key = Item
    Cells(Counter, 1).Value = Item
    Cells(Counter, 2).Value = Json(Key)
    Counter = Counter + 1
Next
    
End Sub

Sub Summary(Ribbon As IRibbonControl)
URL = "https://api.covid19api.com/summary"
MainRun (URL)

'' Handle Response
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

End Sub

Sub CountryStats(Ribbon As IRibbonControl)
Country = ActiveCell.Value

'' Ensure active Cell is in column 3 for country slug
If ActiveCell.Column <> 3 Then
    MsgBox "You must select Country from Column 3 to run detail"
    Exit Sub
End If

URL = "https://api.covid19api.com/dayone/country/" & Country
MainRun (URL)

'' Handle Response
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
End Sub
