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
	Response.Write("<Script>alert('������д������˶Ժ�������д��');history.back();</Script>")
	Response.End
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�༭����ɹ�</title>
<link href="style.css" rel="stylesheet" type="text/css">
<meta http-equiv="refresh" content="2;URL=admin_sortlist.asp">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td align="center"><font color="#FFFFFF">�༭����ɹ�</font></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#EEFEE0">2���Ӻ󷵻ط����б�ҳ�档</td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>