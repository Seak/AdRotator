<%
If Session("AdRotator_Login") <> True Then
	Response.Redirect "index.asp"
	Response.End
End If
%>