<%@ Language="VBScript" %>
<!-- include the database connection file -->
<!--#include file="db_connection.asp"-->

<html>
<head>
    <meta charset="UTF-8">
    <title>Received Messages</title>
    <!-- link the stylesheet -->
    <link rel="stylesheet" type="text/css" href="card.css">
</head>
<body>

    <!-- page title -->
    <h2 style="text-align:center;">Received Messages</h2>

<%
' define variables for recordset and alignment
Dim rs, i, alignmentClass

' retrieve all messages from the database ordered by date (ascending)
Set rs = conn.Execute("SELECT * FROM Messages ORDER BY CreatedDate ASC")

' counter to alternate alignment
i = 0

' loop through each message
Do Until rs.EOF
    i = i + 1

    ' alternate message alignment: even = right, odd = left
    If i Mod 2 = 0 Then
        alignmentClass = "message-card right"
    Else
        alignmentClass = "message-card left"
    End If
%>

    <!-- message card -->
    <div class="<%=alignmentClass%>">
        <div class="message-header">
            <!-- show full name and email -->
            <strong><%=rs("FullName")%></strong> — <%=rs("Email")%>
        </div>
        <div class="message-body">
            <!-- encode message text for HTML safety -->
            <p><%=Server.HTMLEncode(rs("Message"))%></p>
        </div>
        <div class="message-footer">
            <!-- show message creation date -->
            <small><%=rs("CreatedDate")%></small>
        </div>
    </div>

<%
    ' move to the next message
    rs.MoveNext
Loop

' close and clean up objects
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>

    <!-- link to go back to the homepage -->
    <p style="text-align:center;"><a href="index.asp" class="back-link">← Back to Home</a></p>

</body>
</html>
