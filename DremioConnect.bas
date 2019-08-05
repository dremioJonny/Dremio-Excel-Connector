Attribute VB_Name = "DremioConnect"
Option Explicit

Public dremioToken As String
Public DremioClient As New WebClient
Public Response As Object
Public InitialRequest As Boolean

Sub Connect()
    Dim h As String
    Dim pt As String
    Dim usernm  As String
    Dim ps As String
    Dim sslen As String

    h = Sheets("Setup").Range("B1").Value
    pt = Sheets("Setup").Range("B2").Value
    usernm = Sheets("Setup").Range("B3").Value
    ps = Sheets("Setup").Range("B4").Value
    sslen = Sheets("Setup").Range("B5").Value
    
    If IsEmpty(h) = True Then
        MsgBox "Username is empty"
        Exit Sub
    End If
    
    If IsEmpty(pt) = True Then
        MsgBox "Port is empty"
        Exit Sub
    End If
    
    If IsEmpty(usernm) = True Then
        MsgBox "Username is empty"
        Exit Sub
    End If
    
    If IsEmpty(ps) = True Then
        MsgBox "Password is empty"
        Exit Sub
    End If
    
    If IsEmpty(sslen) = True Then
        MsgBox "SSL value not entered. Defaulting to FALSE"
        sslen = "False"
    End If
    
    InitialConnection h, pt, usernm, ps, sslen
End Sub
    
    
Private Sub InitialConnection(h As String, pt As String, u As String, ps As String, ssl As String)
    Dim bodyString As String
    Dim DremioRequest As New WebRequest
    
    DremioClientInit h, pt, ssl
    
    bodyString = "{" & Chr(34) & "userName" & Chr(34) & ": " & Chr(34) & Trim(u) & Chr(34) & "," & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & Trim(ps) & Chr(34) & "}"
    DremioRequest.Body = bodyString
    DremioRequest.Resource = "/apiv2/login"
    DremioRequest.Method = WebMethod.HttpPost
    DremioRequest.Format = WebFormat.Json
    
    Set Response = JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)
    
    AddDocumentVariable ("_dremio" & Response("token"))
    
    If InitialRequest <> True Then
        returnDremioDSs
        InitialRequest = True
    End If
    
End Sub

Private Sub returnDremioDSs()
    Dim DremioRequest As New WebRequest
    Dim QueryString As String
    Dim bodyString As String
    
    DremioRequest.Resource = "api/v3/sql"
    DremioRequest.Method = WebMethod.HttpPost
    DremioRequest.Format = WebFormat.Json
    UseDocumentVariable
    DremioRequest.AddHeader "Authorization", dremioToken
    
    QueryString = "SELECT DISTINCT TABLE_SCHEMA as NAME_SPACE FROM INFORMATION_SCHEMA.\u0022TABLES\u0022 WHERE TABLE_SCHEMA NOT IN (\u0027INFORMATION_SCHEMA\u0027,\u0027sys\u0027,\u0027$scratch\u0027) ORDER BY TABLE_SCHEMA ASC"
    bodyString = "{" & Chr(34) & "sql" & Chr(34) & ": " & Chr(34) & QueryString & Chr(34) & "}"
    DremioRequest.Body = bodyString

    Set Response = JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)
    
    DremioRequest.Method = WebMethod.HttpGet
    DremioRequest.Format = WebFormat.Json
    DremioRequest.Resource = ("/api/v3/job/" & Response("id"))
    
    While JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)("jobState") <> "COMPLETED"
        Application.Wait (Now + TimeValue("0:00:01"))
    Wend
    
    Dim coll As Collection
    
    DremioRequest.Resource = ("/api/v3/job/" & Response("id") & "/results?limit=500")
    Set coll = JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)("rows")
    Dim key As Object
    
    Dim ddoptions As String
    ddoptions = " "
    For Each key In coll
        ddoptions = ddoptions & "," & (key("NAME_SPACE"))
    Next key
    
    Debug.Print (ddoptions)
    
    Dim rDV As Range
    Set rDV = Sheets("Setup").Range("C5")
    
    With rDV.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=ddoptions
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
End Sub
