<!--#include file="conn.asp" -->
<!--#include file="inc_config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ƽ���Ϲ������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">

</head>

<body background="images/back.gif">
<table width="100%" border="0">
  <tr>
    <td align="center"><p><font color="#0000FF" size="5" face="����">��ӭʹ�� 
        ƽ���Ϲ������ϵͳ</font></p>

      <table width="100%" border="0">
        <tr> 
          <td width="25%" rowspan="2">&nbsp;</td>
          <td> <ol>
              <li>��ϵͳ�� <strong>ƽ���Ϲ�����</strong> ��Ա <strong>������</strong> ������ɡ�</li>
              <li>��ϵͳΪ��������(shareware)�ṩ������վ���ʹ�á�</li>
              <li>�Ǿ�������Ȩ��ɣ����ý�֮����ӯ�����ӯ���Ե���ҵ��;��</li>
              <li>Ϊ��Ӧʵ�ʵļ����Ӧ�û������߸Ľ��书�ܡ����ܣ����Խ��б�Ҫ���޸ģ�������ȥ�����ߵİ�Ȩ��ʾ�����ý��޸ĺ�汾�����κε���ҵ��Ϊ�������޸ĺ�İ汾�������ߡ�</li>
              <li>������Ϊ���������û�����ѡ���Ƿ�ʹ�ã���ʹ���г����κ��������ɵ���ʧ���߲����κ����Ρ� </li>
              <li><font color="#999999"><strong>ʹ�ñ�ϵͳ�����Ҫע�ᣬ����Ϊ��֧�������ܹ�������������õĳ���ϣ����һ�����������һ��ע����á����ö���û��Ӳ�Թ涨����λʹ����������������ɡ�</strong></font></li>
            </ol></td>
          <td width="20%" rowspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td><p>��ϵ��ʽ��<br>
              ���ͣ�River.Yin��<br>
              �����ʺţ������������Ҫ�ɣ��ٺ�<br>
              QQ:11265943<br>
              E-Mail:River@pnnic.com<br>
              http://www.pnnic.com/<br>
            </p>
            </td>
        </tr>
      </table>
      <% If Session("AdRotator_Login") <> True Then %>
      <form name="form1" method="post" action="admin_login.asp" target="_top">
        ����Ա�� 
        <input type="text" name="AdmName">
        <br>
        �ܡ��룺 
        <input type="Password" name="AdmPassword">
        <br>
        <input type="submit" name="Submit" value="�ύ">
        ��
<input type="reset" name="Reset" value="����">
      </form>
<% End If %>
      <p><br>
      </p></td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
</body>
</html>
