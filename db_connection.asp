<%
' declare the database connection object
Dim conn

' create an ADO connection object
Set conn = Server.CreateObject("ADODB.Connection")

' open a connection to the Access database
' - using Microsoft ACE OLEDB 12.0 provider
' - Server.MapPath resolves the relative path to an absolute path on the server
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("messagesdb.accdb")
%>
