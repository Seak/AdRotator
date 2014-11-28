<!--#include file="conn.asp"-->
<%
Dim Rs
If Request("cmd") = "update" Then
	ConnectionDatabase
	Conn.Execute("Alter Table [AdList] Add [AdStopShow] int Default 0")
	Conn.Execute("Alter Table [AdList] Add [AdStopClick] int Default 0")
	Conn.Execute("Alter Table [AdList] Add [AdStopDateTime] datetime Default Now()")
	Response.Write("数据结构升级成功！")

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "Select dtmAddDate, AdStopDateTime From [AdList]", Conn, 3, 2
	Do While Not Rs.EOF
		Rs("AdStopDateTime") = Rs("dtmAddDate")
		Rs.Update
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write("<br>数据更新成功！<br>请删除本文件！")
	CloseDatabase
Else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>系统升级</title>

</head>

<body background="images/back.gif">
升级前请先备份数据库！<br>
<a href="?cmd=update">开始升级</a>
</body>
</html>
<% End If %>