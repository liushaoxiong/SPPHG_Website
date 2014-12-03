<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title>在线咨询|<%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
a{ color:#000; font-size:12px;text-decoration:none}
a:hover{ color:#900; text-decoration:underline}
#massage_box{ position:absolute; left:expression((body.clientWidth-550)/2); top:expression((body.clientHeight-250)/2); width:550px; height:340px;filter:dropshadow(color=#666666,offx=3,offy=3,positive=2); z-index:2; visibility:hidden}
#mask{ position:absolute; top:0; left:0; width:expression(body.scrollWidth); height:expression(body.scrollHeight); background:#666; filter:ALPHA(opacity=60); z-index:1; visibility:hidden}
.massage{border:#f45b0d solid; border-width:1 1 3 1; width:95%; height:95%; background:#fff; color:#f45b0d; font-size:12px; line-height:150%}
.header{background:#f45b0d; height:10%; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; padding:8 5 0 5; color:#fff}
</style>
<script language=JavaScript>
function checkdata() {
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
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="8">&nbsp;</td>
    <td width="225" valign="top"><!--#include file="../Web/left.html"--></td>
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="Bor">
      <tr>
        <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_35.jpg">
          <tr>
            <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle">在线留言</td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td align="center" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="60" align="left"><span onClick="mask.style.visibility='visible';massage_box.style.visibility='visible'" style="cursor:hand"><a href="#"><img src="../E_img/zxzx.jpg" width="95" height="32" /></a></span></td>
            </tr>
          </table>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td align="left"><%
Set Rs = Server.CreateObject("ADODB.Recordset")

Sql="select * from E_Message where IfShow = 1 order by ID Desc"

rs.open sql,conn,1,1
dim CurPage,StartPageNum,EndPageNum,p,sizepage,dispage
filename = "/Web/MessageList.asp"'此文件的文件名 
sizepage =10'设置每页记录数 
dispage =10'设置页面上显示多少页 
if rs.eof and rs.bof then%>
                没有找到任何记录!
                <%else 
rs.PageSize = sizepage 
Dim TotalPages 
TotalPages = rs.PageCount 

if Not isnumeric(Request.QueryString("CurPage")) or Request.QueryString("CurPage") = "" Then 
CurPage = 1 
Elseif Cdbl(Request.QueryString("CurPage")) > TotalPages Then 
CurPage = TotalPages 
Else 
CurPage = Cint(Request.QueryString("CurPage")) 
End If 

rs.AbsolutePage=CurPage 
rs.CacheSize = rs.PageSize'设置最大记录数 
Dim Totalcount 
Totalcount =RS.recordcount 

StartPageNum=1 
do while StartPageNum+dispage<=CurPage 
StartPageNum=StartPageNum+dispage 
Loop 

EndPageNum=StartPageNum+dispage-1 
%></td>
            </tr>
          </table>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <%
If EndPageNum>rs.Pagecount then EndPageNum=rs.Pagecount
I=0 
p=RS.PageSize*(Curpage-1) 
do while (Not rs.Eof) and (I<rs.PageSize) 
p=p+1
iRs = I + 1
%>
            <tr>
              <td height="30" align="left" class="Dot T14"><strong><%=Rs("MesName")%></strong>：<%=Rs("MesTitle")%></td>
            </tr>
            <tr>
              <td height="30" align="left" class="Dot Line22" style="padding:8px 0 8px 0"><%=Rs("MesContent")%></td>
            </tr>
            <tr>
              <td height="30" align="left" bgcolor="f1f7ff" class="Line22" style="padding:8px 0 8px 0"><%=Rs("ReContent")%></td>
            </tr>
            <tr>
              <td height="10">&nbsp;</td>
            </tr>
            <%
I=I+1 
rs.MoveNext 
Loop
%>
          </table>
          <table width="95%" border="0" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
            <tr bgcolor="">
              <td width="36%" height="21">页次：<b><font color="#FF0000"><%=CurPage%></font>/<%=TotalPages%></b> 每页<b><%=sizepage%></b> 总记录数<b><%=rs.recordcount%>　</b></td>
              <td width="64%" background="../image/0000999.jpg" bgcolor=""><div class="scott">分页：
                <%if CurPage>dispage then%>
                      <a href="<%=filename%>?CurPage=1"><font title="首页">&lt;&lt;</font></a>
                      <%end if%>
                      <%if CurPage>dispage then%>
                      <a href="<%=filename%>?CurPage=<%=StartPageNum-1%>"><font title="上<%=dispage%>">&lt;</font></a>
                      <%end if 
                  For I=StartPageNum to EndPageNum 
                  if I<>CurPage then %>
                      <a href="<%=filename%>?CurPage=<%=I%>"><%=I%></a>
                      <% else %>
                      <span class="fya"><%=I%></span></font>
                      <% end if 
                  Next %>
                      <% if EndPageNum<RS.Pagecount then %>
                      <a href="<%=filename%>?CurPage=<%=EndPageNum+1%>"><font title="下<%=dispage%>">&gt;</font></a>
                      <%end if 
                  if CurPage<TotalPages then%>
                      <a href="<%=filename%>?CurPage=<%=TotalPages%>"><font title="尾页">&gt;&gt;</font></a>
                      <%end if%>
                      <%end if%>
              </div></td>
            </tr>
          </table></td>
      </tr>
    </table></td>
    <td width="8">&nbsp;</td>
  </tr>
</table>
<!--#include file="../Web/Foot.html" -->
<div id="massage_box"><div class="massage">
<div class="header" onmousedown=MDown(massage_box)><div style="display:inline; width:150px; position:absolute"><span style="padding-top:8px; font-size:14px"><strong>发布留言</strong></span></div>
<span onClick="massage_box.style.visibility='hidden'; mask.style.visibility='hidden'" style="float:right; display:inline; cursor:hand; font-size:18px"><strong>×</strong></span></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<form action="Save.asp?SaveType=MessageAdd" method="post" name="formWrite" id="formWrite" onSubmit="return checkdata()">
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