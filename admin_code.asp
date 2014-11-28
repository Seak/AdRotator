<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<!--#include file="inc_check.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>参考代码</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="100%" border="0">
  <tr>
    <td><div align="center">
        <p><font color="#0000FF" size="5" face="隶书">参考代码</font></p>
        <p>调用方法：<br>
          1.随机显示所有的广告： 
          <textarea name="textarea" cols="55" rows="2" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"><script src="<%= Path %>show.asp"></script></textarea>
          <br>
          2.固定显示某一条广告： 
          <textarea name="textarea2" cols="55" rows="2" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"><script src="<%= Path %>show.asp?id=n"></script></textarea>
          <br>
          3.随机显示某分类广告： 
          <textarea name="textarea3" cols="55" rows="2" style="BACKGROUND-COLOR: #EEFEE0; BORDER: 1 SOLID; FONT-SIZE: 9pt"><script src="<%= Path %>show.asp?sort=n"></script></textarea>
        </p>
      </div></td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>