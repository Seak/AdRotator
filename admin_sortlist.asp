<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����б�</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td><font color="#FFFFFF">���</font></td>
    <td><font color="#FFFFFF">�����б�</font></td>
    <td><div align="center"><font color="#FFFFFF">����</font></div></td>
  </tr>
  <tr> 
    <td width="15%" bgcolor="#EEFEE0">0</td>
    <td width="60%" bgcolor="#EEFEE0">�������</td>
    <td width="25%" bgcolor="#EEFEE0"><font color="#000000"><s>�༭</s></font>��<font color="#000000"><s>ɾ��</s></font> 
    </td>
  </tr>
<%
ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select * From [AdSort]", Conn, 3, 1
Do While Not Rs.EOF
%>
  <tr> 
    <td width="15%" bgcolor="#EEFEE0"><%= Rs("ID") %></td>
    <td width="60%" bgcolor="#EEFEE0"><%= Rs("strSortName") %></td>
    <td width="25%" bgcolor="#EEFEE0"><a href="admin_sortedit.asp?id=<%= Rs("ID") %>" target="_self"><font color="#000000">�༭</font></a>��<a href="admin_sortdeling.asp?id=<%= Rs("ID") %>" target="_self"><font color="#000000">ɾ��</font></a> 
    </td>
  </tr>
<%
	Rs.MoveNext       
Loop
Rs.Close
Set Rs = Nothing

CloseDatabase
%>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>