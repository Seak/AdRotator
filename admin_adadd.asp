<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ӹ��</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <form name="form1" method="post" action="admin_adaddsave.asp">
    <tr> 
      <td><div align="center"><font color="#FFFFFF">��ӹ��</font></div></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#EEFEE0">���ࣺ 
        <select name="intSortID" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt">
<%
ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select ID, strSortName From [AdSort]", Conn, 3, 1
Do While Not Rs.EOF
%>
          <option value="<%= Rs("ID") %>"><%= Rs("strSortName") %></option>
<%
	Rs.MoveNext       
Loop
Rs.Close
Set Rs = Nothing

CloseDatabase
%>
        </select> </td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">���ƣ� 
        <input type="text" name="strAdName" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">��ַ�� 
        <input type="text" name="strAdURL" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">ͼƬ�� 
        <input type="text" name="strAdBanner" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">˵���� 
        <input type="text" name="strAdExplain" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
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
</body>
</html>