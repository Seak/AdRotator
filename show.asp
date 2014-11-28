<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<%
ID = Request.QueryString("id")
Sort = Request.QueryString("sort")

If ID <> "" Then Where = "Where ID = " & ID
If Sort <> "" Then Where = "Where intSortID = " & Sort
If Where <> "" Then
	Where = Where & " And "
Else
	Where = Where & "Where "
End If

Where = Where & "(AdStopShow = 0 Or AdStopShow > intAdShow) And (AdStopClick = 0 Or AdStopClick > intAdClick) And (AdStopDateTime = dtmAddDate Or AdStopDateTime > dtmAddDate)"

ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select ID, strAdName, strAdBanner, strAdExplain, intAdShow, intAdClick, dtmAddDate, AdStopShow, AdStopClick, AdStopDateTime From [AdList]" & Where, Conn, 3, 2
If Not(Rs.EOF) Then
	Rs.MoveFirst
	Randomize
	Rs.Move Int(rs.RecordCount * Rnd)

	Rs("intAdShow") = Rs("intAdShow") + 1
	Rs.Update
%>
<% If Right(Rs("strAdBanner"), 3) <> "swf" Then %>
document.write('<a href="<%= Path %>redirect.asp?id=<%= Rs("ID") %>" title="')
document.writeln('名称：<%= Rs("strAdName") %>')
document.writeln('显示：<%= Rs("intAdShow") %>')
document.writeln('点击：<%= Rs("intAdClick") %>')
document.writeln('加入：<%= Rs("dtmAddDate") %>')
document.write('说明：<%= Rs("strAdExplain") %>')
document.write('">')
document.write('<img src="<%= Rs("strAdBanner") %>" width="468" height="60" border="0">')
document.write('</a>')
<% Else %>
document.write('<embed src="<%= Rs("strAdBanner") %>" width="468" height="60"></embed>')
<% End If %>
<%
End If
Rs.Close
Set Rs = Nothing

CloseDatabase
%>