<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�༭����</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<%
ID = Request.QueryString("id")

ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select ID, strSortName, dtmAddDate From [AdSort] Where ID = " & ID, Conn, 3, 1
%>
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <form name="form1" method="post" action="admin_sorteditsave.asp?id=<%= ID %>">
    <tr> 
      <td><div align="center"><font color="#FFFFFF">�༭����</font></div></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0"> ���ƣ� 
        <input type="text" name="strSortName" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25" value="<%= Rs("strSortName") %>"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0"> ʱ�䣺 
        <input type="text" name="dtmAddDate" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25" value="<%= Rs("dtmAddDate") %>"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0"> <input type="submit" name="Submit" value="�ύ" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt">
        �� 
        <input type="reset" name="Reset" value="����" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"> 
      </td>
    </tr>
  </form>
</table>
<!--#include file="inc_footer.asp" -->
<%
Rs.Close
Set Rs = Nothing

CloseDatabase
%>
</body>
</html>