<%
Dim Rs
Dim ID
Dim AdmName, AdmPassword, Path, RecordPerPage, Version
Dim NewAdmName, NewAdmPassword, NewPath, NewRecordPerPage
Dim strAdName, strAdURL, strAdBanner, strAdExplain, intAdShow, intAdClick, dtmAddDate
Dim absPageNum, TotalPages, absRecordNum

ConnectionDatabase

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select * From [System] Where ID = 1", Conn, 1, 1

AdmName = Rs("AdmName")
AdmPassword = Rs("AdmPassword")
Path = Rs("Path")
RecordPerPage = Rs("RecordPerPage")

Rs.Close
Set Rs = Nothing

CloseDatabase

Version = "1.0"
%>