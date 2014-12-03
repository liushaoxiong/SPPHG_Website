<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<%
ClassID = int(request("ID"))
Set rs = Server.CreateObject("ADODB.Recordset")
if ClassID="" or not IsNumeric(ClassID) then
  sql="select top 1 * from E_About"
else
  sql="select * from E_About where ID="&ClassID
end if  
rs.open sql,conn,1,3

TiTle=rs("TiTle")	
Content=rs("Content")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title><%=Title%>|<%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="../Web/Top.html" -->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="8">&nbsp;</td>
    <td width="225" valign="top"><!--#include file="../Web/left.html"--></td>
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="Bor">
      <tr>
        <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_35.jpg">
          <tr>
            <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle"><%=Title%></td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="20" align="left"></td>
          </tr>
          <tr>
            <td align="left" class="T14 line25"><%=Content%></td>
          </tr>
          <tr>
            <td height="30">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>
    </td>
    <td width="8">&nbsp;</td>
  </tr>
</table>
<!--#include file="../Web/Foot.html" -->
</body>
</html>
<%rs.close
set rs=nothing
%>