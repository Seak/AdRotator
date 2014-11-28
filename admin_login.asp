<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<%
If Request.Form("AdmName") = AdmName And Request.Form("AdmPassword") = AdmPassword Then
	Session("AdRotator_Login") = True
	Response.Redirect("admin_index.asp")
Else
	Response.Redirect("index.asp")
End If
%>