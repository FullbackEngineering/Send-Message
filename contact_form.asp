<%@ Language="VBScript" %>
<!--#include file="db_connection.asp"-->
<html>
<head>
    <meta charset="UTF-8">
    <title>Send Message</title>
    <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>

    <h2>Send New Message</h2>

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim fullname, email, message, sql
    fullname = Request.Form("fullname")
    email = Request.Form("email")
    message = Request.Form("message")

    sql = "INSERT INTO Messages (FullName, Email, Message) VALUES ('" & _
    Replace(fullname, "'", "''") & "', '" & _
    Replace(email, "'", "''") & "', '" & _
    Replace(message, "'", "''") & "')"
    conn.Execute sql
    Response.Write "<p class='success-msg'>Your message has been sent successfully. <a href='index.asp'>Return to home page</a></p>"
End If
%>

    <form method="post" action="contact_form.asp" class="form-container">
        <input type="text" name="fullname" placeholder="Full Name" required>
        <input type="email" name="email" placeholder="Email" required>
        <textarea name="message" placeholder="Write your message..." rows="5" required></textarea>
        <input type="submit" value="Send">
    </form>

    <p><a href="index.asp" class="back-link">← Home Page</a></p>

</body>
</html>
