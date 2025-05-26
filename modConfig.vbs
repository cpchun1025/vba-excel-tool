' This module has private access only!
Option Explicit

Private Type EnvSettings
    ApiUrl As String
    ApiToken As String
End Type

Private Envs As Object

' Called once to set up environment settings
Public Sub InitEnvs()
    Set Envs = CreateObject("Scripting.Dictionary")
    
    Dim dev As EnvSettings, uat As EnvSettings, prod As EnvSettings
    
    dev.ApiUrl = "https://dev-api.example.com/api/save"
    dev.ApiToken = "dev-token-here"
    
    uat.ApiUrl = "https://uat-api.example.com/api/save"
    uat.ApiToken = "uat-token-here"
    
    prod.ApiUrl = "https://prod-api.example.com/api/save"
    prod.ApiToken = "prod-token-here"
    
    Envs.Add "DEV", dev
    Envs.Add "UAT", uat
    Envs.Add "PROD", prod
End Sub

' Returns the EnvSettings object for the given environment
Public Function GetEnvSettings(env As String) As EnvSettings
    If Envs Is Nothing Then InitEnvs
    GetEnvSettings = Envs(env)
End Function

Dim Envs As Object ' Scripting.Dictionary

Sub InitEnvs()
    Set Envs = CreateObject("Scripting.Dictionary")

    Dim dev As New EnvSettings
    dev.ApiUrl = "https://dev-api.example.com/api/save"
    dev.ApiToken = "dev-token-here"

    Dim uat As New EnvSettings
    uat.ApiUrl = "https://uat-api.example.com/api/save"
    uat.ApiToken = "uat-token-here"

    Envs.Add "DEV", dev
    Envs.Add "UAT", uat
End Sub

Function GetEnvSettings(env As String) As EnvSettings
    If Envs Is Nothing Then InitEnvs
    Set GetEnvSettings = Envs(env)
End Function