Attribute VB_Name = "LoadData"
Sub loaddata()
    
    Dim ctlog As String
    Dim tabName As String
    Dim DremioRequest As New WebRequest
    Dim QueryString As String
    Dim bodyString As String
    Dim h As String
    Dim pt As String
    Dim ssl As String
    
    ctlog = Sheets("Setup").Range("C5").Value
    tabName = Sheets("Setup").Range("D5").Value
    h = Sheets("Setup").Range("B1").Value
    pt = Sheets("Setup").Range("B2").Value
    ssl = Sheets("Setup").Range("B5").Value
    
    DremioClientInit h, pt, ssl
    DremioRequest.Resource = "api/v3/sql"
    DremioRequest.Method = WebMethod.HttpPost
    DremioRequest.Format = WebFormat.Json
    UseDocumentVariable
    DremioRequest.AddHeader "Authorization", dremioToken
    
    ctlog = Replace(ctlog, ".", "\u0022.\u0022")
    tabName = Replace(tabName, ".", "\u0022.\u0022")
    QueryString = "SELECT * FROM \u0022" & ctlog & "\u0022.\u0022" & tabName & "\u0022 ORDER BY 1"
    bodyString = "{" & Chr(34) & "sql" & Chr(34) & ": " & Chr(34) & QueryString & Chr(34) & "}"
    DremioRequest.Body = bodyString

    Set Response = JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)
    
    DremioRequest.Method = WebMethod.HttpGet
    DremioRequest.Format = WebFormat.Json
    DremioRequest.Resource = ("/api/v3/job/" & Response("id"))
    
    While JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)("jobState") <> "COMPLETED"
        Application.Wait (Now + TimeValue("0:00:01"))
    Wend
    
    Dim coll As New Collection
    Dim noruns As Integer
    
    DremioRequest.Resource = ("/api/v3/job/" & Response("id") & "/results?limit=500")
    Set coll = JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)("rows")
    noruns = Round_Up((JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)("rowCount")) / 500)
    Dim key As Object
    Dim row As Double
    Dim roworig As Integer
    Dim runnumber As Double
    runnumber = 1
    
    row = 7
    roworig = 7
    
    Dim ddoptions As String
    For Each key In coll
        ddoptions = PrintToWS(key, row, roworig, 0)
    Next key
    
    While runnumber <= noruns:
        DremioRequest.Resource = ("/api/v3/job/" & Response("id") & "/results?offset=" & (runnumber * 500) & "&limit=500")
        Dim collext As Collection
        Set collext = JsonConverter.ParseJson(DremioClient.Execute(DremioRequest).Content)("rows")
        For Each key In collext
            ddoptions = PrintToWS(key, row, roworig, 0)
        Next key
        runnumber = runnumber + 1
    Wend
End Sub


