<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<%
intSortID = Request.Form("intSortID")
strAdName = Server.HTMLEncode(Request.Form("strAdName"))
strAdURL = Server.HTMLEncode(Request.Form("strAdURL"))
strAdBanner = Server.HTMLEncode(Request.Form("strAdBanner"))
strAdExplain = Server.HTMLEncode(Request.Form("strAdExplain"))
dtmAddDate = Now()
AdStopShow = Server.HTMLEncode(Request.Form("AdStopShow"))
AdStopClick = Server.HTMLEncode(Request.Form("AdStopClick"))
AdStopDateTime = Server.HTMLEncode(Request.Form("AdStopDateTime"))
If AdStopDateTime = "0" Then AdStopDateTime = dtmAddDate

If strAdName <> "" And strAdURL <> "" And strAdBanner <> "" And strAdExplain <> "" And AdStopShow <> "" And AdStopClick <> "" And AdStopDateTime <> "" Then
	ConnectionDatabase
	Conn.Execute("Insert Into [AdList] (intSortID, strAdName, strAdURL, strAdBanner, strAdExplain, dtmAddDate, AdStopShow, AdStopClick, AdStopDateTime) Values ('"& intSortID &"', '"& strAdName &"', '"& strAdURL &"', '"& strAdBanner &"', '"& strAdExplain &"', '"& dtmAddDate &"', '"& AdStopShow &"', '"& AdStopClick &"', '"& AdStopDateTime &"')")
	CloseDatabase
Else
	Response.Write("<Script>alert('������д������˶Ժ�������д��');history.back();</Script>")
	Response.End
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ӹ��ɹ�</title>
<link href="style.css" rel="stylesheet" type="text/css">
<meta http-equiv="refresh" content="2;URL=admin_adlist.asp">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td align="center"><font color="#FFFFFF">��ӹ��ɹ�</font></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#EEFEE0">2���Ӻ󷵻ع���б�ҳ�档</td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>