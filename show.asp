<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<%
ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select ID, strAdName, strAdBanner, strAdExplain, intAdShow, intAdClick, dtmAddDate From [AdList]", Conn, 3, 2
If Not(Rs.EOF) Then
	Rs.MoveFirst
	Randomize
	Rs.Move Int(rs.RecordCount * Rnd)

	Rs("intAdShow") = Rs("intAdShow") + 1
	Rs.Update
%>
document.write('<a href="<%= Path %>redirect.asp?id=<%= Rs("ID") %>" title="')
document.writeln('名称：<%= Rs("strAdName") %>')
document.writeln('显示：<%= Rs("intAdShow") %>')
document.writeln('点击：<%= Rs("intAdClick") %>')
document.writeln('加入：<%= Rs("dtmAddDate") %>')
document.write('说明：<%= Rs("strAdExplain") %>')
document.write('">')
document.write('<img src="<%= Rs("strAdBanner") %>" border="0">')
document.write('</a>')
<%
End If
Rs.Close
Set Rs = Nothing

CloseDatabase
%>