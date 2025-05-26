' EnvSettings.cls
Option Explicit

Private pApiUrl As String
Private pApiToken As String

Public Property Get ApiUrl() As String
    ApiUrl = pApiUrl
End Property
Public Property Let ApiUrl(Value As String)
    pApiUrl = Value
End Property

Public Property Get ApiToken() As String
    ApiToken = pApiToken
End Property
Public Property Let ApiToken(Value As String)
    pApiToken = Value
End Property