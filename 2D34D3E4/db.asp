<%
Dim conn
' Microsoft SQL Server ODBC Driver
Set conn = Server.CreateObject("ADODB.Connection")
connString = connString & "DRIVER={SQL Server};"
connString = connString & "DATABASE=vendor_interview;"
connString = connString & "SERVER=TONYMAC\SQLEXPRESS;"

' open the connection here using: Connstr, "username", "password"
conn.Open connString, "tony", "tonymac168"

%>