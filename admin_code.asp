<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ο�����</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="100%" border="0">
  <tr>
    <td><div align="center">
        <p><font color="#0000FF" size="5" face="����">�ο�����</font></p>
        <p>���÷�����<br>
          1.�����ʾ���еĹ�棺 
          <textarea name="textarea" cols="55" rows="2" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"><script src="<%= Path %>show.asp"></script></textarea>
          <br>
          2.�̶���ʾĳһ����棺 
          <textarea name="textarea2" cols="55" rows="2" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"><script src="<%= Path %>show.asp?id=n"></script></textarea>
          <br>
          3.�����ʾĳ�����棺 
          <textarea name="textarea3" cols="55" rows="2" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"><script src="<%= Path %>show.asp?sort=n"></script></textarea>
        </p>
      </div></td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>