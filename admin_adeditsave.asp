<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<%
ID = CInt(Request.QueryString("id"))
intSortID = Request.Form("intSortID")
strAdName = Request.Form("strAdName")
strAdURL = Request.Form("strAdURL")
strAdBanner = Request.Form("strAdBanner")
strAdExplain = Request.Form("strAdExplain")
intAdShow = Request.Form("intAdShow")
intAdClick = Request.Form("intAdClick")
dtmAddDate = Request.Form("dtmAddDate")
AdStopShow = Request.Form("AdStopShow")
AdStopClick = Request.Form("AdStopClick")
AdStopDateTime = Request.Form("AdStopDateTime")
If AdStopDateTime = "0" Then AdStopDateTime = dtmAddDate

If ID <> "" And strAdName <> "" And strAdURL <> "" And strAdBanner <> "" And strAdExplain <> "" And intAdShow <> "" And intAdClick <> "" And dtmAddDate <> "" And AdStopShow <> "" And AdStopClick <> "" And AdStopDateTime <> "" Then
	ConnectionDatabase
	Conn.Execute("Update [AdList] Set intSortID = '"& intSortID &"', strAdName = '"& strAdName &"', strAdURL = '"& strAdURL &"', strAdBanner = '"& strAdBanner &"', strAdExplain = '"& strAdExplain &"', intAdShow = '"& intAdShow &"', intAdClick = '"& intAdClick &"', dtmAddDate = '"& dtmAddDate &"', AdStopShow = '"& AdStopShow &"', AdStopClick = '"& AdStopClick &"', AdStopDateTime = '"& AdStopDateTime &"' Where ID = " & ID)
	CloseDatabase
Else
	Response.Write("<Script>alert('������д������˶Ժ�������д��');history.back();</Script>")
	Response.End
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�༭���ɹ�</title>
<link href="style.css" rel="stylesheet" type="text/css">
<meta http-equiv="refresh" content="2;URL=admin_adlist.asp">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td align="center"><font color="#FFFFFF">�༭���ɹ�</font></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#EEFEE0">2���Ӻ󷵻ع���б�ҳ�档</td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>