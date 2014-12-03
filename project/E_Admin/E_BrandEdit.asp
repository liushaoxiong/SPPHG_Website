<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<!--#include file="../E_Inc/E_BClassSClass.asp"-->
<%
'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|6,")=0 then 
  response.write "<script language='JavaScript'>alert('您无权操作此页面，请返回！');" & "history.back()" & "</script>"
  response.end
end if
end if

Default=3
ClassMenu = 4
Menu=""
Menu1=1
ID=FunIfInt(request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from E_Brand where  id="&ID
Rs.Open sql,conn,1,1
if Rs.eof and Rs.bof then
response.Write("没有记录")
else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>兰州宜天网络管理后台</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="E_Admin.js"></script>
<script language=JavaScript>
function checkdata() {
if (editForm.BrandName.value.length < 1) {
alert("\请填写品牌名称!!")
editForm.BrandName.focus()
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
              <td class="Wryh T14 P2"><strong>&nbsp;修改品牌</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" background="Img/Add_banner.jpg"><!--#include file="BrandMenu.html"--></td>
              </tr>
            </table>
             <div id="Empty"></div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=BrandEdit&ID=<%=ID%>" onSubmit="return checkdata()">
                 <tr>
                   <td width="100" height="40" align="center" class="Wryh T14">品牌名称</td>
                   <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                         <td background="Img/Input_03.jpg"><input name="BrandName" class="Titleinput" value="<%=Rs("BrandName")%>"></td>
                         <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                       </tr>
                   </table></td>
                 </tr>
                 <tr>
                   <td width="100" height="100" align="center" class="Wryh T14">品牌简介</td>
                   <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td width="8"><img src="Img/input1_01.jpg" width="8" height="86"></td>
                         <td background="Img/Input1_03.jpg"><textarea name="Content" rows="4" class="Title1input"><%If Rs("Content") <> "" Then%><%=changechr2(Rs("Content"))%><%Else%><%=Rs("Content")%><%End if%></textarea></td>
                         <td width="8"><img src="Img/input1_05.jpg" width="8" height="86"></td>
                       </tr>
                   </table></td>
                 </tr>
                 <tr>
                   <td width="100" height="40" align="center" class="Wryh T14">品牌Logo</td>
                   <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
                             <tr>
                               <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                               <td background="Img/Input_03.jpg"><input name="Pictures" class="Titleinput"  value="<%=Rs("Pictures")%>"></td>
                               <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                             </tr>
                         </table></td>
                         <td width="20%"><a href="javaScript:OpenScript('UpFileForm.asp?Result=Pictures',460,180)"><img src="Img/Upload.jpg" width="91" height="30" border="0" align="absmiddle"></a></td>
                         <td width="10%">&nbsp;</td>
                       </tr>
                   </table></td>
                 </tr>

                 <tr>
                   <td width="100" height="50" align="center">&nbsp;</td>
                   <td><input type="image" src="Img/Tijiao.jpg"></td>
                 </tr>
               </form>
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