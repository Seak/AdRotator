<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�༭���</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<%
ID = Request.QueryString("id")

ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select ID, intSortID, strAdName, strAdURL, strAdBanner, strAdExplain, intAdShow, intAdClick, dtmAddDate From [AdList] Where ID = " & ID, Conn, 3, 1
%>
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <form name="form1" method="post" action="admin_adeditsave.asp?id=<%= ID %>">
    <tr> 
      <td><div align="center"><font color="#FFFFFF">�༭���</font></div></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">���ࣺ 
        <select name="intSortID" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt">
<%
Set RsSort = Server.CreateObject("ADODB.Recordset")
RsSort.Open "Select * From [AdSort]", Conn, 3, 1
Do While Not RsSort.EOF
%>
          <option value="<%= RsSort("ID") %>"<% If RsSort("ID") = Rs("intSortID") Then Response.Write(" Selected") %>><%= RsSort("strSortName") %></option>
<%
	RsSort.MoveNext       
Loop
RsSort.Close
Set RsSort = Nothing
%>
        </select>
        </td>
    </tr>
    <tr>
      <td align="center" bgcolor="#EEFEE0">���ƣ� 
        <input name="strAdName" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strAdName") %>" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">��ַ�� 
        <input name="strAdURL" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strAdURL") %>" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">ͼƬ�� 
        <input name="strAdBanner" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strAdBanner") %>" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">˵���� 
        <input name="strAdExplain" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strAdExplain") %>" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">ʱ�䣺 
        <input name="dtmAddDate" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("dtmAddDate") %>" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">��ʾ�� 
        <input name="intAdShow" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("intAdShow") %>" size="7">
        ������� 
        <input name="intAdClick" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("intAdClick") %>" size="7"></td>
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