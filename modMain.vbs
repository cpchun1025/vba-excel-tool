Dim gUser As String, gPwd As String, gEnv As String

Sub ShowLoginAndSetEnv()
    frmLogin.Show
    gUser = frmLogin.txtUser.Text
    gPwd = frmLogin.txtPwd.Text
    gEnv = frmLogin.cmbEnv.Value
End Sub

Sub SendData_Batch()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 1. User login and env selection
    ShowLoginAndSetEnv

    ' 2. Get environment settings
    Dim envSettings As EnvSettings
    envSettings = modConfig.GetEnvSettings(gEnv)
    Dim apiUrl As String
    apiUrl = envSettings.ApiUrl
    Dim apiToken As String
    apiToken = envSettings.ApiToken
    
    ' 3. Read template from shared drive
    Dim templatePath As String
    templatePath = "Y:\templates\batch_template.json" ' Or build path dynamically per sheet
    Dim template As String
    template = ReadTemplateFromFile(templatePath)

    ' 4. [As before] Build JSON payload using VBA-JSON, substitute {{Rows}}, etc.
    ' (Use previous batch code, but with the template string loaded from disk)

    ' 5. Send to API
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    With http
        .Open "POST", apiUrl, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiToken
        ' If API expects username/pwd as headers, add here:
        '.setRequestHeader "X-User", gUser
        '.setRequestHeader "X-Pwd", gPwd
        .send finalJson
        ' Handle response/logging as before
    End With
End Sub