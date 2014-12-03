<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<!--#include file="../E_Inc/E_BClassSClass.asp"-->
<%
'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|5,")=0 then 
  response.write "<script language='JavaScript'>alert('您无权操作此页面，请返回！');" & "history.back()" & "</script>"
  response.end
end if
end if

Default=3
ClassMenu = 3
BClassCount=Conn.execute("Select count(Id) from BClass where IfShow = 0")(0)
ID=FunIfInt(request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from SClass where  id="&ID
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
<style type="text/css">
#massage_box{ position:absolute; left:expression((body.clientWidth-400)/2); top:200px; width:600px; height:260px;filter:dropshadow(color=#666666,offx=3,offy=3,positive=2); z-index:2; visibility:hidden}
#mask{ position:absolute; top:0; left:0; width:expression(body.scrollWidth); height:expression(body.scrollHeight); background:#666; filter:ALPHA(opacity=60); z-index:1; visibility:hidden}
.massage{border:#333 solid; border-width:1 1 3 1; width:95%; height:95%; background:#fff; color:#40bffd; font-size:12px; line-height:150%}
.header{background:#f3f3f3; height:20px; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; padding:10px 5 10 5; color:#333}

#massage_box1{ position:absolute; left:expression((body.clientWidth-400)/2); top:220px; width:400px; height:260px;filter:dropshadow(color=#666666,offx=3,offy=3,positive=2); z-index:2; visibility:hidden}
#mask1{ position:absolute; top:0; left:0; width:expression(body.scrollWidth); height:expression(body.scrollHeight); background:#666; filter:ALPHA(opacity=60); z-index:1; visibility:hidden}
.massage1{border:#333 solid; border-width:1 1 3 1; width:95%; height:95%; background:#fff; color:#40bffd; font-size:12px; line-height:150%}
.header1{background:#f3f3f3; height:20px; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; padding:10px 5 10 5; color:#333}
</style>
<script language="javascript">
function ifmsgbox(message,url) {
	     if ( confirm(message) ) {
		 window.location.href=url
		 }
	}
</script>
<script language=JavaScript>
function checkdata() {
if (editForm.BClassName.value.length < 1) {
alert("\请填写分类名称!!")
editForm.BClassName.focus()
return false;
}
return true;
}
</script>

<script language=JavaScript>
function checkdata1() {
if (editForm1.UpId.value == '0') {
alert("\分类不能为空，请返回选择!!")
editForm1.UpId.focus()
return false;
}
if (editForm1.SClassName.value.length < 1) {
alert("\请填写分类名称!!")
editForm1.SClassName.focus()
return false;
}
return true;
}
</script>
</head>
<body>
<!--#include file="Top.html"-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" valign="top" bgcolor="f2f2f2">
    <!--#include file="E_SysMenu.asp"-->
    </td>
    <td width="205" valign="top" background="Img/C_bg.jpg">
    <!--#include file="E_ProMenu.asp"-->
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
              <td class="Wryh T14 P2"><strong>&nbsp;<%=FunBClassName(Rs("UpId"))%>&nbsp;-&nbsp;小类管理</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" background="Img/Add_banner.jpg"><!--#include file="ProMenu.html"--></td>
              </tr>
            </table>
            <div id="Empty"></div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=SClassEdit&ID=<%=ID%>&UpId=<%=Rs("UpId")%>">
                 <tr>
                   <td width="100" height="40" align="center" class="Wryh T14">分类名称</td>
                   <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                         <td background="Img/Input_03.jpg"><input name="SClassName" class="Titleinput" value="<%=Rs("SClassName")%>"></td>
                         <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                       </tr>
                   </table></td>
                 </tr>

                 <tr>
                   <td width="100" height="50">&nbsp;</td>
                   <td><input type="image" src="Img/Tijiao.jpg"></td>
                 </tr>
               </form>
          </table>
          </td>
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
<!--#include file="FootWindows.asp"-->
</body>
</html>
<%
end if
Rs.close
set Rs=nothing
%>