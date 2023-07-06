<%@ Language="VBScript" CodePage = 65001 %>
<pre><%

Server.ScriptTimeout = 120

' Test urls
httpGET "https://www.example.com/"
httpGET "https://www.canada.ca"
httpGET "https://forces.gc.ca"
httpGET "http://intranet.mil.ca"

' Debug GET request
Function httpGET(url)
    On Error Resume Next

    Dim xmlhttp: Set xmlhttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

    If Err.Number <> 0 Then
        Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
        Err.clear
    End If

    Response.Write "Created " & TypeName(xmlhttp) & "<br>"

    setTimeouts xmlhttp, 5000, 5000, 10000, 10000
    xmlhttp.open "GET", url, False
    setOption xmlhttp, 4, 13056
    showOptions(xmlhttp)

    If Err.Number <> 0 Then
        Response.Write "Open<br>    Error " & Err.Number & " " & Err.Description & "<br>"
    Else
        Response.Write "Opened<br>"
        xmlhttp.send
    
        If Err.Number <> 0 Then
            Response.Write "Send<br>    Error " & Err.Number & " " & Replace(Err.Description, vbNewLine, "") & "<br>"
        Else
            Response.Write "Sent<br>"
        End If
    End If
   
    showResponse xmlhttp

    Response.write "<br><br>"
    Set xmlhttp = Nothing
    Err.clear
End Function

' Set and print timeouts
Function setTimeouts(byref xmlhttp, a, b, c, d)
    Response.Write "  Timeouts (ms)<br>"
    Response.Write "    Resove: " & a & "<br>"
    Response.Write "    Connect: " & b & "<br>"
    Response.Write "    Send: " & c & "<br>"
    Response.Write "    Recieve: " & d & "<br>"
    
    xmlhttp.setTimeouts a, b, c, d
End Function

' Set option
Function setOption(byref xmlhttp, i, v)
    On Error Resume Next

    If TypeName(xmlhttp) = "WinHttpRequest" Then
        xmlhttp.Option(i) = xmlhttp.Option(i) Or v
    Else
        xmlhttp.setOption i, xmlhttp.getOption(i) Or v
    End If

    Err.clear
End Function

' Print options
Function showOptions(byref xmlhttp)
    On Error Resume Next
    Response.Write("  Options<br>")

    Dim names(10)
    names(1) = "1-URL"
    names(2) = "2-URL_CODEPAGE"
    names(3) = "3-ESCAPE_PERCENT_IN_URL"
    names(4) = "4-IGNORE_SERVER_SSL_CERT_ERROR_FLAGS"
    names(5) = "5-SELECT_CLIENT_SSL_CERT"
    names(6) = "6-ENABLE_REDIRECTS"
    names(7) = "7-ENABLE_HTTP_REDIRECT"
    names(8) = "8-ENABLE_AUTHENTICATION"
    names(9) = "9-SERVER_CERT_IGNORE_FLAGS"

    If TypeName(xmlhttp) = "WinHttpRequest" Then
        For i = 1 To 9
            Response.Write "    " & names(i) & ": " & xmlhttp.Option(i) & "<br>"
        Next
    Else
        For i = 1 To 9
            Response.Write "    " & names(i) & ": " & xmlhttp.getOption(i) & "<br>"
        Next
    End If

    Err.clear
End Function

' Print response data
Function showResponse(byref xmlhttp)
    Response.Write "Response<br>"
    Response.Write "  Status<br>    " & xmlhttp.status & " " & xmlhttp.statusText & "<br>"
    Response.Write "  Headers<br>"

    Dim header
    For Each header In Split(xmlhttp.getAllResponseHeaders(), vbCrLf)
        If Len(header) > 0 Then
            Response.Write "    " & header & "<br>"
        End If
    Next
End Function

%></pre>
