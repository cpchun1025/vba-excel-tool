Sub GetTokenWithWindowsAuth()
    Dim http As Object
    Dim url As String
    Dim response As String

    url = "https://sso.company.com/gettoken"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetAutoLogonPolicy 0 ' 0 = Always, 1 = OnlyIfBypassProxy, 2 = Never

    ' This will use the currently logged-in Windows user's credentials
    http.Send

    If http.Status = 200 Then
        response = http.ResponseText
        MsgBox "Token: " & response
    Else
        MsgBox "Failed: " & http.Status & " - " & http.StatusText
    End If
End Sub