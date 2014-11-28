<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告列表</title>
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
    <td> <div align="left"><font color="#FFFFFF">第<%= absPageNum %>/<%= TotalPages %>页　本页<%= RecordPerPage %>条　共<%= rs.RecordCount %>条 
        </font> </div></td>
    <td><div align="right"><font color="#FFFFFF"> 
        <% If Not absPageNum = 1 Then %>
        [<a href="admin_adlist.asp?page=1" target="_self">首页</a>][<a href="admin_adlist.asp?page=<%= absPageNum - 1 %>" target="_self">上一页</a>] 
        <% End If %>
        <% If Not absPageNum = TotalPages Then %>
        [<a href="admin_adlist.asp?page=<%= absPageNum + 1 %>" target="_self">下一页</a>][<a href="admin_adlist.asp?page=<%= TotalPages %>" target="_self">尾页</a>] 
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
    <td width="150"><font color="#FFFFFF">序号：<b><font color="#FF0000"><%= Rs("ID") %></font></b></font></td>
    <td width="150"><font color="#FFFFFF">分类：<%= Rs("intSortID") %></font></td>
    <td colspan="2"><div align="right"><font color="#FFFFFF">管理：<a href="admin_adedit.asp?id=<%= Rs("ID") %>" target="_self">编辑</a>　<a href="admin_addeling.asp?id=<%= Rs("ID") %>">删除</a></font></div></td>
  </tr>
  <tr> 
    <td colspan="4" bgcolor="#EEFEE0">名称：<%= Rs("strAdName") %></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#EEFEE0">时间：<%= Rs("dtmAddDate") %><% If Rs("AdStopDateTime") <> Rs("dtmAddDate") Then Response.Write(" / <b>" & Rs("AdStopDateTime") & "</b>") %>
    </td>
    <td width="90" bgcolor="#EEFEE0">显示：<%= Rs("intAdShow") %> 
      <% If Rs("AdStopShow") <> 0 Then Response.Write("/<b>" & Rs("AdStopShow") & "</b>") %>
    </td>
    <td width="90" bgcolor="#EEFEE0">点击：<%= Rs("intAdClick") %> 
      <% If Rs("AdStopClick") <> 0 Then Response.Write("/<b>" & Rs("AdStopClick") & "</b>") %>
    </td>
  </tr>
  <tr> 
    <td colspan="4" bgcolor="#EEFEE0">地址：<%= Rs("strAdURL") %></td>
  </tr>
  <tr> 
    <td colspan="4" bgcolor="#EEFEE0">图片：<%= Rs("strAdBanner") %></td>
  </tr>
  <tr> 
    <td colspan="4" bgcolor="#EEFEE0">说明：<%= Rs("strAdExplain") %></td>
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
    <td> <div align="left"><font color="#FFFFFF">第<%= absPageNum %>/<%= TotalPages %>页　本页<%= RecordPerPage %>条　共<%= rs.RecordCount %>条 
        </font> </div></td>
    <td><div align="right"><font color="#FFFFFF"> 
        <% If Not absPageNum = 1 Then %>
        [<a href="admin_adlist.asp?page=1" target="_self">首页</a>][<a href="admin_adlist.asp?page=<%= absPageNum - 1 %>" target="_self">上一页</a>] 
        <% End If %>
        <% If Not absPageNum = TotalPages Then %>
        [<a href="admin_adlist.asp?page=<%= absPageNum + 1 %>" target="_self">下一页</a>][<a href="admin_adlist.asp?page=<%= TotalPages %>" target="_self">尾页</a>] 
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