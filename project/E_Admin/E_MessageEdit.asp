<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
Default=4
ClassMenu = 2
'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|14,")=0 then 
  response.write "<script language='JavaScript'>alert('您无权操作此页面，请返回！');" & "history.back()" & "</script>"
  response.end
end if
end if

ID=FunIfInt(request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from E_Message where  id="&ID
Rs.Open sql,conn,1,1
if Rs.eof and Rs.bof then
response.Write("没有记录")
else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网络管理后台</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="E_Admin.js"></script>
</head>
<body>
<!--#include file="Top.html"-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" valign="top" bgcolor="f2f2f2">
    <!--#include file="E_SysMenu.asp"-->
    </td>
    <td width="205" valign="top" background="Img/C_bg.jpg"><!--#include file="OptionMenu.html"--></td>
    <td width="20"></td>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FAFAFA">
        <tr>
          <td align="center" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="Img/menu_12.jpg">
            <tr>
              <td width="47"><img src="Img/menu_10.jpg" width="47" height="55"></td>
              <td class="Wryh T14 P2"><strong>&nbsp;留言管理</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
          <div id="Empty"></div>
            <table width="95%" border="0" cellpadding="0" cellspacing="1" bgcolor="#3E4B6B">
            <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=MesEdit&ID=<%=ID%>">
            <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">留言姓名：</td>
                <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="90" class="T14 Wryh">&nbsp;<%=Rs("MesName")%></td>
    <td class="T14 Wryh">&nbsp;<strong>性别：</strong><%=Rs("MesSex")%></td>
  </tr>
</table>
</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">留言标题：</td>
                <td bgcolor="#FFFFFF" class="T14 Wryh">&nbsp;<%=Rs("MesTitle")%></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">留言内容：</td>
                <td bgcolor="#FFFFFF" class="T14 Wryh Line25">&nbsp;<%=Rs("MesContent")%></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">审核状态：</td>
                <td bgcolor="#FFFFFF" class="T14 Wryh Line25">&nbsp;<input name="IfShow" type="checkbox" value="1" <%if Rs("IfShow") = 1 Then%> checked<%End if%>></td>
              </tr>
              <tr>
                <td width="100" height="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">回复内容：</td>
                <td bgcolor="#FFFFFF" class="T14 Wryh Line25"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="15" align="right"><img src="Img/input1_01.jpg" width="8" height="86"></td>
                    <td background="Img/Input1_03.jpg"><textarea name="ReContent" rows="4" class="Title1input"><%If Rs("ReContent") <> "" Then%><%=changechr2(Rs("ReContent"))%><%Else%><%=Rs("ReContent")%><%End if%></textarea></td>
                    <td width="8"><img src="Img/input1_05.jpg" width="8" height="86"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="50" align="center" bgcolor="#FFFFFF">&nbsp;</td>
                <td bgcolor="#FFFFFF">&nbsp;<input type="image" src="Img/Tijiao.jpg"></td>
              </tr>
              </form>
          </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="30">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="95%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14"><img src="Img/Centet_10.jpg" width="14" height="13"></td>
          <td background="Img/Centet_11.jpg"><img src="Img/Centet_11.jpg"></td>
          <td width="7"><img src="Img/Centet_13.jpg" width="7" height="13"></td>
        </tr>
      </table>
      <table width="95%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="40">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Foot.html"-->
</body>
</html>
<%
end if
Rs.close
set Rs=nothing
%>