<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title>友情链接|<%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
body{overflow-x:hidden;}
a{ color:#000; font-size:12px;text-decoration:none}
a:hover{ color:#900; text-decoration:underline}
#massage_box{ position:absolute; left:expression((body.clientWidth-550)/2); top:expression((body.clientHeight-200)/2); width:550px; height:340px;filter:dropshadow(color=#666666,offx=3,offy=3,positive=2); z-index:2; visibility:hidden}
#mask{ position:absolute; top:0; left:0; width:expression(body.scrollWidth); height:expression(body.scrollHeight); background:#666; filter:ALPHA(opacity=60); z-index:1; visibility:hidden}
.massage{border:#ffe0fd solid; border-width:1 1 3 1; width:95%; height:95%; background:#fff; color:#ffe0fd; font-size:12px; line-height:150%}
.header{background:#9a0000; height:10%; font-family:"微软雅黑"; font-size:12px; padding:8 5 0 5; color:#fff}
</style>
<script language=JavaScript>
function checkdata1() {
if (formWrite.MesName.value.length < 1) {
alert("\请填写姓名!!")
formWrite.MesName.focus()
return false;
}
if (formWrite.MesTitle.value.length < 1) {
alert("\请填写标题!!")
formWrite.MesTitle.focus()
return false;
}
if (formWrite.MesContent.value.length < 1) {
alert("\请填写内容!!")
formWrite.MesContent.focus()
return false;
}
return true;
}
</script>
</head>
<body>
<!--#include file="../Web/Top.html" -->
<table width="1190" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="112" valign="top"><!--#include file="../Web/left.html"--></td>
    <td align="center" valign="top"><table width="96%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="right">您当前的位置：<a href="../">首页</a>&nbsp;>&nbsp;友情链接</td>
      </tr>
    </table>
        <table width="96%" border="0" cellpadding="0" cellspacing="0" background="../E_img/main_04.jpg">
          <tr>
            <td width="105" valign="top" class="ClassName1" align="center">友情链接</td>
            <td height="18"><img src="../E_img/main_04.jpg"></td>
          </tr>
        </table>
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30">&nbsp;</td>
          </tr>
        </table>
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td class="Line25">
        <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select * from E_FriendLink where LinkType = 1 order by id Asc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
<a href="<%=Rs("LinkUrl")%>" target="_blank"><%=Rs("LinkName")%></a>　
<% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>     
</td>
          </tr>
          <tr>
            <td class="Dot">&nbsp;</td>
          </tr>
                    <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
        <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select * from E_FriendLink where LinkType = 2 and Pictures <> '' order by id Asc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
<td><a href="<%=Rs("LinkUrl")%>" target="_blank" title="<%=Rs("LinkName")%>"><img src="<%=Rs("Pictures")%>" width="110" height="39"></a>　</td>
<% 
If 	iRs mod 6 = 0 Then
Response.Write("</tr><tr><td height=20></td></tr><tr>")
end if	
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
  </tr>
</table>

            </td>
          </tr>
        </table>
</td>
    <td width="230" align="right" valign="top"><!--#include file="../Web/Right.html"--></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html"-->
<div id="massage_box"><div class="massage">
<div class="header" onmousedown=MDown(massage_box)><div style="display:inline; width:150px; position:absolute"><span style="padding-top:8px; font-size:14px"><strong>发布留言</strong></span></div>
<span onClick="massage_box.style.visibility='hidden'; mask.style.visibility='hidden'" style="float:right; display:inline; cursor:hand; font-size:18px"><strong>×</strong></span></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<form action="Save.asp?SaveType=MessageAdd" method="post" name="formWrite" id="formWrite" onSubmit="return checkdata1()">
  <tr>
    <td width="80" height="30" align="center"><strong>姓名：</strong></td>
    <td><input name="MesName" type="text" class="input"  id="MesName" size="16"/></td>
  </tr>
  <tr>
    <td width="80" height="30" align="center"><strong>标题：</strong></td>
    <td><input name="MesTitle" type="text" class="input"  id="MesTitle"size="56"/></td>
  </tr>
  <tr>
    <td width="80" height="30" align="center"><strong>内容：</strong></td>
    <td><textarea name="MesContent" cols="55" rows="5" class="MesContent"></textarea></td>
  </tr>
  <tr>
    <td width="80" height="30" align="center"><strong>验证：</strong></td>
    <td><input name="VerifyCode" type="text" class="input" style="width:98px;" size="4" maxlength="4"/>　<span onClick="document.getElementById('PhotoSN').src='../E_Inc/E_VerifyCode.asp?'+Math.random()" style="cursor: hand"><img src="../E_Inc/E_VerifyCode.asp" name="PhotoSN" width="60" height="22" id="PhotoSN" style="border:1px solid #7f9db9;" /></span></td>
  </tr>
  <tr>
    <td height="40" align="center">&nbsp;</td>
    <td><input name="Submit" type="submit" class="fya" value=" 保 存 " />
    &nbsp;&nbsp;<input name="Reset" type="reset" class="fya" value=" 重 置 " /></td>
  </tr>
  </form>
</table>

</div></div>
<div id="mask"></div>
</body>
</html>