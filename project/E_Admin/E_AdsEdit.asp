<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
Default=4
ClassMenu = 3

'========�ж��Ƿ���й���Ȩ��
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|13,")=0 then 
  response.write "<script language='JavaScript'>alert('����Ȩ������ҳ�棬�뷵�أ�');" & "history.back()" & "</script>"
  response.end
end if
end if

ID=FunIfInt(request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from E_Ads where  id="&ID
Rs.Open sql,conn,1,1
if Rs.eof and Rs.bof then
response.Write("û�м�¼")
else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������̨</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="E_Admin.js"></script>
<script language=JavaScript>
function checkdata() {
if (editForm.AdsTitle.value == "") {
alert("\������Ʋ���Ϊ�գ��뷵������!!")
editForm.AdsTitle.focus()
return false;
}
//if (editForm.AdsContent.value == "") {
//alert("\����ַ����Ϊ�գ��뷵��!!")
//editForm.AdsContent.focus()
//return false;
//}
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
    <td width="205" valign="top" background="Img/C_bg.jpg"><!--#include file="OptionMenu.html"--></td>
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
              <td class="Wryh T14 P2"><strong>&nbsp;�༭���</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" background="Img/Add_banner.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="20">&nbsp;</td>
                    <td width="103" height="27" background="Img/qb1.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="30">&nbsp;</td>
                          <td align="center" class="white Wryh T14"><a href="E_AdsMain.asp"><strong>����б�</strong></a></td>
                        </tr>
                    </table></td>
                    <td width="10">&nbsp;</td>
                    <td width="103" background="Img/tj.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="30">&nbsp;</td>
                          <td align="center" class="white Wryh T14"><a href="E_AdsAdd.asp"><strong>��ӹ��</strong></a></td>
                        </tr>
                    </table></td>
                    <td width="10">&nbsp;</td>
                    <td width="103" background="Img/Back.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="30">&nbsp;</td>
                          <td align="center" class="white Wryh T14"><a href="Javascript:history.back();"><strong>�����ϼ�</strong></a></td>
                        </tr>
                    </table></td>
                    <td>&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
            </table>
             <div id="Empty"></div>
             <table width="100%" border="0" cellspacing="0" cellpadding="0">
               <tr>
                 <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                   <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=AdsEdit&ID=<%=ID%>" onSubmit="return checkdata()">
                     <tr>
                       <td width="100" height="40" align="center" class="Wryh T14">�������</td>
                       <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                           <tr>
                             <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                             <td background="Img/Input_03.jpg"><input name="AdsTitle" class="Titleinput" value="<%=Rs("AdsTitle")%>"></td>
                             <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                           </tr>
                       </table></td>
                     </tr>
                     <tr>
                       <td width="100" height="40" align="center" class="Wryh T14">���ӵ�ַ</td>
                       <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                           <tr>
                             <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                             <td background="Img/Input_03.jpg"><input name="AdsLink" class="Titleinput" value="<%=Rs("AdsLink")%>"></td>
                             <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                           </tr>
                       </table></td>
                     </tr>
                     <tr>
                       <td width="100" height="40" align="center" class="Wryh T14">�������</td>
                       <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                           <tr>
                             <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
                                 <tr>
                                   <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                                   <td background="Img/Input_03.jpg"><input name="AdsContent" class="Titleinput" value="<%=Rs("AdsContent")%>"></td>
                                   <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                                 </tr>
                             </table></td>
                             <td width="20%"><a href="javaScript:OpenScript('UpFileForm.asp?Result=AdsContent',460,180)"><img src="Img/Upload.jpg" width="91" height="30" border="0" align="absmiddle"></a></td>
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
                 <td width="200" align="center">
<%If Rs("AdsContent") <> "" Then%>
<%
AdImgType = GetExtendName(Rs("AdsContent"))
	select Case  AdImgType
	Case "swf"
	response.write "<object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 width=180>"
	response.write "<param name=movie value="&Rs("AdsContent")&">"
	response.write "<param name=quality value=high>"
	response.write "<embed src="&Rs("AdsContent")&" quality=high pluginspage=http://www.macromedia.com/go/getflashplayer type=application/x-shockwave-flash width=180 height=40></embed>"
	response.write "</object>"
	Case Else
		response.write "<img src="&Rs("AdsContent")&" width=180 height=40>"
	END select				 
%>
<%End if%>
</td>
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