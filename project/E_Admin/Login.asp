<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<script language="javascript">
function CheckForm()
{
	if(document.Login.LoginName.value=="")
	{
		alert("�������û�����");
		document.Login.LoginName.focus();
		return false;
	}
	if(document.Login.LoginPassword.value == "")
	{
		alert("���������룡");
		document.Login.LoginPassword.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("������������֤�룡");
       document.Login.CheckCode.focus();
       return(false);
    }
}
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������̨</title>
<link href="Img/Css.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="80">&nbsp;</td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="370" valign="top" bgcolor="#4882D6"><table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="10">&nbsp;</td>
      </tr>
    </table>
      <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="670" valign="top"><img src="Img/lbg.jpg" width="423" height="348"></td>
          <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
              <td height="15"></td>
            </tr>
            <tr>
              <td height="50" class="Wryh T24">�����½</td>
            </tr>
            <tr>
              <td background="Img/Dot.jpg"></td>
            </tr>
            <tr>
              <td height="25">&nbsp;</td>
            </tr>
          </table>
            <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <form name="Login" action="CheckLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
            <tr>
              <td height="55" valign="top" class="inputbg"><input name="UserName" type="text" onFocus="if(this.value=='�����������ʺ�!') {this.value='';}this.style.color='#333333';" onBlur="if(this.value=='') {this.value='�����������ʺ�!';this.style.color='#333333';}"  class="Logininput"></td>
            </tr>
            <tr>
              <td height="65" valign="top" class="inputbg1"><input name="Passwords" type="password" class="Logininput"></td>
            </tr>
            <tr>
              <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><input type="image"  src="Img/Login.jpg"></td>
    </tr>
</table>
</td>
            </tr>
            <tr>
              <td height="30">&nbsp;</td>
            </tr>
            </form>
          </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="945" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50"></td>
  </tr>
</table>
<table width="945" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="80" height="40"><img src="Img/Login_32.jpg" width="73" height="61"></td>
    <td valign="top"><table width="90%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="25" class="Title">��վά��</td>
      </tr>
      <tr>
        <td height="25" class="Textarea">&nbsp;</td>
      </tr>
    </table></td>
    <td width="80"><img src="Img/Login_29.jpg" width="77" height="63"></td>
    <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="25" class="Title">��վ�Ż�</td>
      </tr>
      <tr>
        <td height="25" class="Textarea">&nbsp;</td>
      </tr>
    </table></td>
    <td width="80"><img src="Img/Login_26.jpg" width="77" height="64"></td>
    <td width="230" align="right"><table width="90%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="25" class="Title">��վ�İ�</td>
      </tr>
      <tr>
        <td height="25" class="Textarea">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="945" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"></td>
  </tr>
  <tr>
    <td height="1" bgcolor="dadada"></td>
  </tr>
  <tr>
    <td height="20"></td>
  </tr>
  <tr>
    <td height="20" class="Textarea">��Ȩ���У�Copyright &copy;2007-<%=year(now)%>�����绰��13609317525��QQ:334297092</td>
  </tr>
</table>
</body>
</html>
