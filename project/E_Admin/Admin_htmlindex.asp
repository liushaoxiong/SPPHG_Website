<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<!--#include file="E_htmlconfig.asp"-->
<%
Default=6
ClassMenu = 1

'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|9,") = 0 then 
  response.write "<script language='JavaScript'>alert('您无权操作此页面，请返回！');" & "history.back()" & "</script>"
  response.end
end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网络管理后台</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="Top.html"-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" valign="top" bgcolor="f2f2f2">
    <!--#include file="E_SysMenu.asp"-->
    </td>
    <td width="205" valign="top" background="Img/C_bg.jpg">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="Img/jtsc.jpg" width="205" height="40" /></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
        <td <%If ClassMenu = 1 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_HtmlMain.asp"><img src="Img/C2.png" align="absmiddle">　生成列表</a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr>
     
    </table>
    </td>
    <td width="20"></td>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FAFAFA">
        <tr>
          <td height="400" align="center" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="Img/menu_12.jpg">
            <tr>
              <td width="47"><img src="Img/menu_10.jpg" width="47" height="55"></td>
              <td class="Wryh T14 P2"><strong>&nbsp;静态生成</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
          <div id="Empty"></div>
          <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="300"><table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td style="border-bottom: #ccc 1px solid; border-top: #ccc 1px solid; border-left: #ccc 1px solid; border-right: #ccc 1px solid"><img src="Img/Load.jpg" width="16" height="16" id="bar_img" name="bar_img" align="absmiddle"></td>
        </tr>
      </table></td>
    <td><span id="bar_txt2" name="bar_txt2" style="font-size:12px; color:red;"></span><span id="bar_txt1" name="bar_txt1" style="font-size:12px">0</span><span style="font-size:12px">%</span></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50"><a href=# onClick="javascript:history.back();">请返回</a></td>
  </tr>
</table>

<%

call htmll("/","/","/index.html","Web/Default.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((1/10)*300)&";bar_txt1.innerHTML=""成功生成首页。完成比例" & formatnumber(1/10*100) & """;</script>"
Response.Flush


Response.Write "<script>bar_img.width=300;bar_txt2.innerHTML='';bar_txt1.innerHTML=""成功生成相关静态文件。完成比例：100"";</script>"
Response.End
%></td>
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