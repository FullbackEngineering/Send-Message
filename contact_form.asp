<%@ Language="VBScript" %>
<!-- include the database connection file -->
<!--#include file="db_connection.asp"-->

<html>
<head>
    <meta charset="UTF-8">
    <title>Send Message</title>
    <!-- link to the external CSS file -->
    <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>

    <!-- page heading -->
    <h2>Send New Message</h2>

<%
' check if the request method is POST (form submitted)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' declare and retrieve form inputs
    Dim fullname, email, message, sql
    fullname = Request.Form("fullname")
    email = Request.Form("email")
    message = Request.Form("message")

    ' build SQL query with escaped single quotes to prevent errors
    sql = "INSERT INTO Messages (FullName, Email, Message) VALUES ('" & _
    Replace(fullname, "'", "''") & "', '" & _
    Replace(email, "'", "''") & "', '" & _
    Replace(message, "'", "''") & "')"

    ' execute the insert query
    conn.Execute sql

    ' display confirmation message
    Response.Write "<p class='success-msg'>Your message has been sent successfully. <a href='index.asp'>Return to home page</a></p>"
End If
%>

    <!-- message form -->
    <form method="post" action="contact_form.asp" class="form-container">
        <!-- full name input -->
        <input type="text" name="fullname" placeholder="Full Name" required>

        <!-- email input -->
        <input type="email" name="email" placeholder="Email" required>

        <!-- message textarea -->
        <textarea name="message" placeholder="Write your message..." rows="5" required></textarea>

        <!-- submit button -->
        <input type="submit" value="Send">
    </form>

    <!-- back to home link -->
    <p><a href="index.asp" class="back-link">← Home Page</a></p>

</body>
</html>
