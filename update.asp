<!--#include file="conn.asp"-->
<%
Dim Rs
If Request("cmd") = "update" Then
	ConnectionDatabase
	Conn.Execute("Alter Table [AdList] Add [AdStopShow] int Default 0")
	Conn.Execute("Alter Table [AdList] Add [AdStopClick] int Default 0")
	Conn.Execute("Alter Table [AdList] Add [AdStopDateTime] datetime Default Now()")
	Response.Write("���ݽṹ�����ɹ���")

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "Select dtmAddDate, AdStopDateTime From [AdList]", Conn, 3, 2
	Do While Not Rs.EOF
		Rs("AdStopDateTime") = Rs("dtmAddDate")
		Rs.Update
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write("<br>���ݸ��³ɹ���<br>��ɾ�����ļ���")
	CloseDatabase
Else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳ����</title>

</head>

<body background="images/back.gif">
����ǰ���ȱ������ݿ⣡<br>
<a href="?cmd=update">��ʼ����</a>
</body>
</html>
<% End If %>