<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<%
ID = Request.QueryString("id")

ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select ID, strAdURL, intAdClick From [AdList] Where ID =" & ID, Conn, 3, 2

Rs("intAdClick") = Rs("intAdClick") + 1
Rs.Update

Response.Redirect Rs("strAdURL")

Rs.Close
Set Rs = Nothing

CloseDatabase
%>