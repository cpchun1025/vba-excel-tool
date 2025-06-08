Sub SendTradesToAPI_Async()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim trades As Collection
    Set trades = New Collection

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            Dim trade As Object
            Set trade = CreateObject("Scripting.Dictionary")
            trade("symbol") = ws.Cells(i, 1).Value
            trade("price") = ws.Cells(i, 2).Value
            trade("quantity") = ws.Cells(i, 3).Value
            trade("td") = Format(ws.Cells(i, 4).Value, "yyyy-mm-dd")
            trade("vd") = Format(ws.Cells(i, 5).Value, "yyyy-mm-dd")
            trade("remarks") = ws.Cells(i, 6).Value
            trade("condition") = ws.Cells(i, 7).Value
            trades.Add trade
        End If
    Next i

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "trades", trades
    
    Dim tradesJson As String
    tradesJson = ConvertToJson(payload)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "http://localhost:8000/process-trades/"

    http.Open "POST", url, True
    http.setRequestHeader "Content-Type", "application/json"
    http.send tradesJson

    MsgBox "Trades sent! You can keep working. Please wait for results..."

    ' Poll for completion, allow UI interaction
    Do While http.readyState <> 4
        DoEvents
    Loop

    If http.Status = 200 Then
        Dim response As String
        response = http.responseText

        Dim json As Object
        Set json = ParseJson(response)

        Dim result As Variant
        Dim rowOffset As Long
        rowOffset = 2

        For Each result In json("results")
            ws.Cells(rowOffset, 8).Value = result("tradeid")
            ws.Cells(rowOffset, 9).Value = result("error")
            rowOffset = rowOffset + 1
        Next result

        MsgBox "API processed and wrote results!"
    Else
        MsgBox "API call failed: " & http.Status & " - " & http.statusText
    End If
End Sub
