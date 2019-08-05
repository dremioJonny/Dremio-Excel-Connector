Attribute VB_Name = "InitializeDremioClient"
Option Explicit

Sub DremioClientInit(h As String, pt As String, ssl As String)
    If UCase(ssl) = "TRUE" Then
        DremioClient.BaseUrl = "https://" & Trim(h) & ":" & Trim(pt) & "/"
    Else
        DremioClient.BaseUrl = "http://" & Trim(h) & ":" & Trim(pt) & "/"
        DremioClient.Insecure = True
    End If
End Sub
