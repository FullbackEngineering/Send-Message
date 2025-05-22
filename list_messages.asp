Da<%@ Language="VBScript" %>
<!--#include file="db_connection.asp"-->
<html>
<head>
    <meta charset="UTF-8">
    <title>Received Messages</title>
    <link rel="stylesheet" type="text/css" href="card.css">
</head>
<body>

    <h2 style="text-align:center;">Received Messages</h2>

<%
Dim rs, i, alignmentClass
Set rs = conn.Execute("SELECT * FROM Messages ORDER BY CreatedDate ASC")
i = 0

Do Until rs.EOF
    i = i + 1
    If i Mod 2 = 0 Then
        alignmentClass = "message-card right"
    Else
        alignmentClass = "message-card left"
    End If
%>

    <div class="<%=alignmentClass%>">
        <div class="message-header">
            <strong><%=rs("FullName")%></strong> — <%=rs("Email")%>
        </div>
        <div class="message-body">
            <p><%=Server.HTMLEncode(rs("Message"))%></p>
        </div>
        <div class="message-footer">
            <small><%=rs("CreatedDate")%></small>
        </div>
    </div>

<%
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>

    <p style="text-align:center;"><a href="index.asp" class="back-link">← Back to Home</a></p>

</body>
</html>
