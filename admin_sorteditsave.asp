<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<%
ID = CInt(Request.QueryString("id"))
strSortName = Server.HTMLEncode(Request.Form("strSortName"))
dtmAddDate = Server.HTMLEncode(Request.Form("dtmAddDate"))

If strSortName <> "" And dtmAddDate <> "" Then
	ConnectionDatabase
	Conn.Execute("Update [AdSort] Set strSortName = '"& strSortName &"', dtmAddDate = '"& dtmAddDate &"' Where ID = " & ID)
	CloseDatabase
Else
	Response.Write("<Script>alert('数据填写有误，请核对后重新填写！');history.back();</Script>")
	Response.End
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>编辑分类成功</title>
<link href="style.css" rel="stylesheet" type="text/css">
<meta http-equiv="refresh" content="2;URL=admin_sortlist.asp">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td align="center"><font color="#FFFFFF">编辑分类成功</font></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#EEFEE0">2秒钟后返回分类列表页面。</td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>