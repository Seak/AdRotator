<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳ����</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <form name="form1" method="post" action="admin_systemsave.asp">
    <tr> 
      <td colspan="2"><div align="center"><font color="#FFFFFF">ϵͳ����</font></div></td>
    </tr>
    <tr> 
      <td width="164" align="center" bgcolor="#EEFEE0">����Ա�� 
        <input name="AdmName" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= AdmName %>" size="14"></td>
      <td width="123" align="center" bgcolor="#EEFEE0">��Ӣ�ġ����ֽԿ�</td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">�ܡ��룺 
        <input name="AdmPassword" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= AdmPassword %>" size="14"></td>
      <td align="center" bgcolor="#EEFEE0">����������λ</td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0">�������� 
        <input name="RecordPerPage" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= RecordPerPage %>" size="14"></td>
      <td align="center" bgcolor="#EEFEE0">�б���ÿҳ��ʾ����</td>
    </tr>
<!--    <tr> 
      <td align="center" bgcolor="#EEFEE0">·������ 
        <input name="Path" type="text" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Path %>" size="14"></td>
      <td align="center" bgcolor="#EEFEE0">��ϵͳ�İ�װ·��</td>
    </tr>-->
    <tr> 
      <td colspan="2" align="center" bgcolor="#EEFEE0"> <input type="submit" name="Submit" value="�ύ" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt">
        �� 
        <input type="reset" name="Reset" value="����" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"> 
      </td>
    </tr>
  </form>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>