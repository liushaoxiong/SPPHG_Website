<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<%
ID = int(request("ID"))
Set rs = Server.CreateObject("ADODB.Recordset")
if ID="" or not IsNumeric(ID) then
  sql="select top 1 * from E_About"
else
  sql="select * from E_About where ID="&ID
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
<table width="1190" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="112" valign="top"><!--#include file="../Web/left.html"--></td>
    <td align="center" valign="top"><table width="96%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="right">您当前的位置：<a href="../">首页</a>&nbsp;>&nbsp;<%=Title%></td>
      </tr>
    </table>
 <table width="96%" border="0" cellpadding="0" cellspacing="0" background="../E_img/main_04.jpg">
          <tr>
            <td width="105" valign="top" class="ClassName1" align="center"><%=Title%></td>
            <td height="18"><img src="../E_img/main_04.jpg"></td>
          </tr>
      </table>
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
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
    <td width="230" align="right" valign="top"><!--#include file="../Web/Right.html"--></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html"-->
</body>
</html>
<%rs.close
set rs=nothing
%>