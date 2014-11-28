<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加广告</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <form name="form1" method="post" action="admin_adaddsave.asp">
    <tr> 
      <td><div align="center"><font color="#FFFFFF">添加广告</font></div></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#EEFEE0">分类： 
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
      <td align="center" bgcolor="#EEFEE0">名称： 
        <input type="text" name="strAdName" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">地址： 
        <input type="text" name="strAdURL" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">图片： 
        <input type="text" name="strAdBanner" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">说明： 
        <input type="text" name="strAdExplain" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0"> <input type="submit" name="Submit" value="提交" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt">
        　 
        <input type="reset" name="Reset" value="重置" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"> 
      </td>
    </tr>
  </form>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>