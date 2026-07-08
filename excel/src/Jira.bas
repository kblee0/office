Attribute VB_Name = "Jira"
Option Explicit

Private Const BASE_URL As String = "https://jira.atlassian.com"
Private Const TOKEN As String = "OTEzNjk1OTg6bW8tYXBpLXZpVjVwaWR0WkNuU0FaVnM5bEpMUUl1WQ=="

'Sub jiraSearchExample()
'    Dim res
'
'    res = Jira.search( _
'        "project = ICISTRCO AND ""Task Execution Team"" ~ T-OD", _
'        "issuetype,customfield_10001,summary,assignee,status,resolutiondate,priority,customfield_14000,customfield_14001,customfield_14213", _
'        0, 1000)
'
'    RangeUtil.setArray res, ThisWorkbook.Sheets("REF").Range("A2")
'    RangeUtil.getOffsetRange(ThisWorkbook.Sheets("REF").Range("A2"), UBound(res, 1), 0, 100, UBound(res, 2)).Clear
'End Sub

Private Function encodeUrl(ByVal str As String) As String
  encodeUrl = WorksheetFunction.encodeUrl(str)
End Function

Public Function search(ByVal jql As String, ByVal fields As String, Optional ByVal startAt As Long = 0, Optional ByVal maxResults As Long = 1000)
    Dim json As Object
    Dim body As New Dictionary
    Dim requestBody As String
    Dim fieldNames() As String
    
    fieldNames = Split(fields, ",")
    body("jql") = jql
    body("fields") = fieldNames
    body("maxResults") = maxResults
    
    requestBody = JsonConverter.ConvertToJson(body)
    
    Set json = request("POST", "/rest/api/2/search", Nothing, requestBody)
    
    search = convertIssuesToArray(json, fieldNames)
End Function

Private Function convertIssuesToArray(ByRef jsonObj As Object, ByRef fields() As String) As Variant
    If jsonObj Is Nothing Then
        convertIssuesToArray = Empty
        Exit Function
    End If
    
    Dim total As Long
    Dim result()
    Dim i As Long
    Dim j As Long
    
    total = jsonObj("total")
    
    ReDim result(1 To total, 1 To UBound(fields) + 2)
    
    Dim value
    Dim regex As Object
    Dim isoDate As String
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}.\d{3}[\+\-]\d{4}"
    regex.IgnoreCase = True
    regex.Global = True
    
    For i = 1 To total
        result(i, 1) = jsonObj("issues")(i)("key")
        For j = 0 To UBound(fields)
            If TypeOf jsonObj("issues")(i)("fields")(fields(j)) Is Object  Then
                Set value = jsonObj("issues")(i)("fields")(fields(j))
            Else
                value = jsonObj("issues")(i)("fields")(fields(j))
            End If
            
            Select Case TypeName(value)
            Case "Dictionary"
                If value.Exists("displayName") Then
                    result(i, j + 2) = value("displayName")
                Else
                    result(i, j + 2) = value("name")
                End If
            Case "Collection"
                Dim item
                Dim avalue()
                
                avalue = Array()
                
                For Each item In value
                    ReDim Preserve avalue(0 To UBound(avalue) + 1)
                    
                    If TypeName(item) = "Dictionary" Then
                        avalue(UBound(avalue)) = item("name")
                    Else
                        avalue(UBound(avalue)) = item
                    End If
                Next
                
                quickSort avalue, LBound(avalue), UBound(avalue)
                result(i, j + 2) = Join(avalue, ",")
            Case Else
                If IsNull(value) Then
                    result(i, j + 2) = ""
                ElseIf regex.Test(value) Then
                    isoDate = Left(value, 26) & ":" & Right(value, 2)
                    result(i, j + 2) = JsonConverter.ParseIso(isoDate)
                Else
                    result(i, j + 2) = value
                End If
            End Select
        Next
    Next
    
    Set regex = Nothing

    convertIssuesToArray = result
End Function

Public Function getIssueStatus(ByVal issueKey As String) As String
    Dim json As Object
    Dim params As New Dictionary
    
    params("fields") = "status"
    
    Set json = request("GET", "/rest/api/2/issue/" & issueKey, params)
    
    If json Is Nothing Then
        getIssueStatus = ""
        Exit Function
    End If
    
    getIssueStatus = json("fields")("status")("name")
End Function

Public Function request(ByVal method As String, ByVal url As String, Optional ByVal params As Dictionary = Nothing, Optional ByVal body As String = "") As Object
    Dim jiraurl As String
    Dim http As Object
    Dim json As Object
    Dim paramKey As Variant
    Dim paramVal As String
    
    jiraurl = BASE_URL & url
    
    If Not params Is Nothing Then
        For Each paramKey In params.Keys
            paramVal = params(paramKey)
            
            If jiraurl = (BASE_URL & url) Then
                jiraurl = jiraurl & "?" & encodeUrl(paramKey) & "=" & encodeUrl(paramVal)
            Else
                jiraurl = jiraurl & "&" & encodeUrl(paramKey) & "=" & encodeUrl(paramVal)
            End If
        Next
    End If
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.setTimeouts 2000, 2000, 10000, 30000 ' ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
    http.Option(4) = 13056 ' SSL error skip
    http.Open method, jiraurl, False
    http.SetRequestHeader "Content-Type", "application/json"
    If TOKEN <> "" Then http.SetRequestHeader "Authorization", "Bearer " & TOKEN
    http.SetRequestHeader "User-Agent", "Excel/2024"
    
    On Error Resume Next
    If body = "" Then
        http.send
    Else
        http.send body
    End If
    
    If http.Status <> 200 Then
        Set request = Nothing
    Else
        Set request = JsonConverter.ParseJson(http.ResponseText)
    End If
    
    If Not http Is Nothing Then Set http = Nothing
End Function


Private Sub quickSort(arr As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant
    i = low
    j = high
    pivot = arr((low + high) \ 2)
    While i <= j
        While arr(i) < pivot
            i = i + 1
        Wend
        While arr(j) > pivot
            j = j - 1
        Wend
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend
    If low < j Then quickSort arr, low, j
    If i < high Then quickSort arr, i, high
End Sub

Public Sub addJinkLink()
    Dim r As Range, cell As Range
    Dim font
    
    For Each r In Application.Selection.Areas
        For Each cell In r
            font = cell.font.Name
            ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=BASE_URL & "/browse/" & cell.value, TextToDisplay:=cell.value
            cell.font.Name = font
        Next
    Next
End Sub

