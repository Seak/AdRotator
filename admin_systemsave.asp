<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<%
NewAdmName = Request.Form("AdmName")
NewAdmPassword = Request.Form("AdmPassword")
NewPath = Request.Form("Path")
NewRecordPerPage = Request.Form("RecordPerPage")

If NewAdmName <> "" And NewAdmPassword <> "" And NewPath <> "" And NewRecordPerPage <> "" Then
	ConnectionDatabase
	Conn.Execute("Update [System] Set AdmName = '"& NewAdmName &"', AdmPassword = '"& NewAdmPassword &"', Path = '"& NewPath &"', RecordPerPage = '"& NewRecordPerPage &"' Where ID= 1")
	CloseDatabase
Else
	Response.Write("<Script>alert('������д������˶Ժ�������д��');history.back();</Script>")
	Response.End
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳ���óɹ�</title>
<link href="style.css" rel="stylesheet" type="text/css">
<meta http-equiv="refresh" content="2;URL=admin_adlist.asp">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td align="center"><font color="#FFFFFF">ϵͳ���óɹ�</font></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#EEFEE0">2���Ӻ󷵻ع���б�ҳ�档</td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>