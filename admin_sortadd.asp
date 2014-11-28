<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加分类</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="300" align="center" cellpadding="2" cellspacing="1" bgcolor="#3F8805">
  <form name="form1" method="post" action="admin_sortaddsave.asp">
    <tr> 
      <td><div align="center"><font color="#FFFFFF">添加分类</font></div></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0"> 名称： 
        <input type="text" name="strSortName" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="25"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#EEFEE0"> <input type="submit" name="Submit" value="提交" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt">
        　 
        <input type="reset" name="Reset" value="重置" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"> 
      </td>
    </tr>
  </form>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>