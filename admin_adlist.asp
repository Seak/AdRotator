<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����б�</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<%
ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorLocation = 3
Rs.CacheSize = RecordPerPage
Rs.PageSize = RecordPerPage
Rs.Open "Select * From [AdList] Order By ID", Conn, 3, 1, &H0001

TotalPages = Rs.PageCount
absPageNum = CInt(Request.QueryString("page"))
If Request.QueryString("page") = "" Or absPageNum > TotalPages Then absPageNum = 1
%>
<table width="500" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td> <div align="left"><font color="#FFFFFF">��<%= absPageNum %>/<%= TotalPages %>ҳ����ҳ<%= RecordPerPage %>������<%= rs.RecordCount %>�� 
        </font> </div></td>
    <td><div align="right"><font color="#FFFFFF"> 
        <% If Not absPageNum = 1 Then %>
        [<a href="admin_adlist.asp?page=1" target="_self">��ҳ</a>][<a href="admin_adlist.asp?page=<%= absPageNum - 1 %>" target="_self">��һҳ</a>] 
        <% End If %>
        <% If Not absPageNum = TotalPages Then %>
        [<a href="admin_adlist.asp?page=<%= absPageNum + 1 %>" target="_self">��һҳ</a>][<a href="admin_adlist.asp?page=<%= TotalPages %>" target="_self">βҳ</a>] 
        <% End If %>
        </font></div></td>
  </tr>
</table>
<br>
<%
If Not(Rs.EOF) Then
	Rs.AbsolutePage = absPageNum
	For absRecordNum = 1 To Rs.PageSize
%>
<table width="500" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td width="164"><font color="#FFFFFF">��ţ�<b><font color="#FF0000"><%= Rs("ID") %></font></b></font></td>
    <td width="165"><font color="#FFFFFF">���ࣺ<%= Rs("intSortID") %></font></td>
    <td colspan="2"><font color="#FFFFFF">����<a href="admin_adedit.asp?id=<%= Rs("ID") %>" target="_self">�༭</a>��<a href="admin_addeling.asp?id=<%= Rs("ID") %>">ɾ��</a></font></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#EEFEE0">���ƣ�<%= Rs("strAdName") %></td>
    <td colspan="2" bgcolor="#EEFEE0">ʱ�䣺<%= Rs("dtmAddDate") %></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#EEFEE0">��ַ��<%= Rs("strAdURL") %></td>
    <td width="82" bgcolor="#EEFEE0">��ʾ��<%= Rs("intAdShow") %></td>
    <td width="70" bgcolor="#EEFEE0">�����<%= Rs("intAdClick") %></td>
  </tr>
  <tr> 
    <td colspan="4" bgcolor="#EEFEE0">ͼƬ��<%= Rs("strAdBanner") %></td>
  </tr>
  <tr> 
    <td colspan="4" bgcolor="#EEFEE0">˵����<%= Rs("strAdExplain") %></td>
  </tr>
</table>
<br>
<%
		Rs.MoveNext
		If Rs.EOF Then
			Exit For
		End If
	Next
End If
%>
<table width="500" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <tr> 
    <td> <div align="left"><font color="#FFFFFF">��<%= absPageNum %>/<%= TotalPages %>ҳ����ҳ<%= RecordPerPage %>������<%= rs.RecordCount %>�� 
        </font> </div></td>
    <td><div align="right"><font color="#FFFFFF"> 
        <% If Not absPageNum = 1 Then %>
        [<a href="admin_adlist.asp?page=1" target="_self">��ҳ</a>][<a href="admin_adlist.asp?page=<%= absPageNum - 1 %>" target="_self">��һҳ</a>] 
        <% End If %>
        <% If Not absPageNum = TotalPages Then %>
        [<a href="admin_adlist.asp?page=<%= absPageNum + 1 %>" target="_self">��һҳ</a>][<a href="admin_adlist.asp?page=<%= TotalPages %>" target="_self">βҳ</a>] 
        <% End If %>
        </font></div></td>
  </tr>
</table>
<%
Rs.Close
Set Rs = Nothing

CloseDatabase
%>
<!--#include file="inc_footer.asp" -->
</body>
</html>