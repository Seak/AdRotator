<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>平龙认广告轮显系统</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="100%" border="0">
  <tr>
    <td align="center"><p><font color="#0000FF" size="5" face="隶书">欢迎使用 
        平龙认广告轮显系统</font></p>

      <table width="100%" border="0">
        <tr> 
          <td width="25%" rowspan="2">&nbsp;</td>
          <td> <ol>
              <li>本系统由 <strong>平龙认工作室</strong> 成员 <strong>江海客</strong> 独立完成。</li>
              <li>本系统为共享软体(shareware)提供个人网站免费使用。</li>
              <li>非经作者授权许可，不得将之用于盈利或非盈利性的商业用途。</li>
              <li>为适应实际的计算机应用环境或者改进其功能、性能，可以进行必要的修改；但不得去除作者的版权标示，不得将修改后版本进行任何的商业行为，并将修改后的版本发与作者。</li>
              <li>本软体为免费软件，用户自由选择是否使用，在使用中出现任何问题而造成的损失作者不负任何责任。 </li>
              <li><font color="#999999"><strong>使用本系统无须非要注册，不过为了支持作者能够开发出更多更好的程序，希望大家还是资助作者一点注册费用。费用多少没有硬性规定，各位使用者视自身情况即可。</strong></font></li>
            </ol></td>
          <td width="20%" rowspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td><p>联系方式：<br>
              海客（River.Yin）<br>
              银行帐号：这个……找我要吧，嘿嘿<br>
              QQ:11265943<br>
              E-Mail:River@pnnic.com<br>
              http://www.pnnic.com/<br>
            </p>
            </td>
        </tr>
      </table>
      <% If Session("AdRotator_Login") <> True Then %>
      <form name="form1" method="post" action="admin_login.asp" target="_top">
        管理员： 
        <input type="text" name="AdmName">
        <br>
        密　码： 
        <input type="Password" name="AdmPassword">
        <br>
        <input type="submit" name="Submit" value="提交">
        　
<input type="reset" name="Reset" value="重置">
      </form>
<% End If %>
      <p><br>
      </p></td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>
