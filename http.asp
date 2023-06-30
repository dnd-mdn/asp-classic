<%@ Language="VBScript" CodePage = 65001 %>
<pre><%

Server.ScriptTimeout = 120

' Test urls
httpGET "http://forces.gc.ca"
httpGET "https://forces.gc.ca"
httpGET "http://intranet.mil.ca"
httpGET "https://intranet.mil.ca"

' Debug GET request
Function httpGET(url)
    On Error Resume Next

    Response.Write "GET " & url & "<br>"
    Dim xmlhttp: Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

    If Err.Number <> 0 Then
        Response.Write "> XMLHTTP " & Err.Number & " " & Err.Description & "<br/>"
    Else
        Response.Write "> Created XMLHTTP object<br/>"
        xmlhttp.open "GET", url, False

        If Err.Number <> 0 Then
            Response.Write "> xmlhttp.open " & Err.Number & " " & Err.Description & "<br/>"
        Else
            Response.Write "> xmlhttp.open<br/>"
            xmlhttp.send

            If Err.Number <> 0 Then
                Response.Write "> xmlhttp.send " & Err.Number & " " & Replace(Err.Description, vbNewLine, "") & "<br/>"
            Else
                Response.Write "> xmlhttp.send<br/>"
                Response.Write "> Status " & xmlhttp.status & " " & xmlhttp.statusText & "<br>"
            End If
        End If
    End If

    Response.write "<br>"
    Set xmlhttp = Nothing
    Err.clear
End Function

%></pre>
