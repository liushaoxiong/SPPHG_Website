<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
Default=5
ClassMenu = 2
ClassID = FunIfInt(Request("ClassID"))
'========�ж��Ƿ���й���Ȩ��
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|8,")=0 then 
  response.write "<script language='JavaScript'>alert('����Ȩ������ҳ�棬�뷵�أ�');" & "history.back()" & "</script>"
  response.end
end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������̨</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="E_Admin.js"></script>
<script language=JavaScript>
function checkdata() {
if (editForm.MagName.value.length < 1) {
alert("\����д�û���!!")
editForm.MagName.focus()
return false;
}
if (editForm.MagPass.value.length < 1) {
alert("\����д����!!")
editForm.MagPass.focus()
return false;
}
if( editForm.MagPass.value != editForm.MagPass1.value ) {
alert("\�����������벻һ�£�")
editForm.MagPass.focus()
return false;
}
return true;
}

function changeMagLevel(id)
    {
	if (id == 1) {
		divPopdem2.style.visibility="visible";
		divPopdem.style.visibility="hidden";
	}
	else {
	
		if (id == 3) {
		divPopdem2.style.visibility="visible";
		divPopdem.style.visibility="hidden";
		}
		
		else
		{
		divPopdem2.style.visibility="hidden";
		divPopdem.style.visibility="hidden";
		}
		
	}
	}
function sel() 
{ 
o=document.getElementsByName("MagPopdem") 
for(i=0;i<o.length;i++) 
o[i].checked=event.srcElement.checked 
} 

function PageDisp(url,pagename,widths,heights,scrollbars){
window.open(url,pagename,"width="+widths+",height="+heights+" screenX=50,screenY=50,scrollbars="+scrollbars+",toolbar=no,menubar=no,resizable=no,locationbar=no,hotkeys=no");
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
    <!--#include file="SetSiteMenu.html"-->
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
              <td class="Wryh T14 P2"><strong>&nbsp;����ʺ�</strong></td>
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
                            <td align="center" class="white Wryh T14"><a href="E_AdminUserMain.asp"><strong>�ʺ��б�</strong></a></td>
                          </tr>
                      </table></td>
                      <td width="10">&nbsp;</td>
                      <td width="103" background="Img/tj.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="30">&nbsp;</td>
                            <td align="center" class="white Wryh T14"><a href="E_AdminUserAdd.asp"><strong>����û�</strong></a></td>
                          </tr>
                      </table></td>
                      <td>&nbsp;</td>
                    </tr>
                </table></td>
              </tr>
            </table>
            <div id="Empty"></div> 
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=AdminAdd" onSubmit="return checkdata()">
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">�û���</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagName" class="Titleinput"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>

              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">����</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagPass" type="password" class="Titleinput"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">ȷ������</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagPass1" type="password" class="Titleinput"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">��������</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagEmail" class="Titleinput"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">������</td>
                <td>
                <select name="MagLevel" class="Input" id="MagLevel" onChange="changeMagLevel(document.editForm.MagLevel.value)">
                <option value="0">��������Ա</option>
                 <option value="1">��ͨ����Ա</option></select></td>
              </tr>
              <tr>
                <td width="100" height="50" align="center">&nbsp;</td>
                <td><input type="image" src="Img/Tijiao.jpg"></td>
              </tr>
             
            </table></td>
                <td width="500" valign="top"><div id="divPopdem2" style="position:absolute; visibility:hidden; width: 100%;" align="center">
			<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"><input name="Purview1" type="checkbox" value="|1,">&nbsp;��ҵ��Ϣ��<input name="Purview2" type="checkbox" value="|2,">&nbsp;�༭��ҵ��<input name="Purview3" type="checkbox" value="|3,">&nbsp;���Ź���<input name="Purview7" type="checkbox" value="|7,">&nbsp;ϵͳ����<input name="Purview8" type="checkbox" value="|8,">&nbsp;�ʺŹ���<input name="Purview9" type="checkbox" value="|9,">&nbsp;��̬���ɡ�<input name="Purview10" type="checkbox" value="|10,">&nbsp;ϵͳ����<p><input name="Purview11" type="checkbox" value="|11,">&nbsp;ϵͳ��ԭ��<input name="Purview12" type="checkbox" value="|12,">&nbsp;ϵͳ���ݡ�<input name="Purview13" type="checkbox" value="|13,">&nbsp;������<input name="Purview14" type="checkbox" value="|14,">&nbsp;���Թ���<input name="Purview15" type="checkbox" value="|15,">&nbsp;ǰ̨��Ա</td>
  </tr>
    <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#BCFEFE" class="Wryh T14">&nbsp;<strong>���ŷ���</strong></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#FFFFFF"> 
<% 
Set Rs=Server.Createobject("Adodb.Recordset")
QuerySql="Select * from E_NewsClass order by id Asc"
Rs.open QuerySql,Conn,1,1
For I = 1 to Rs.recordcount
if Rs.EOF then     
Exit For 
end if '����for next ѭ�����ζ�����¼
%><input name="ClassID" type="checkbox" value="|ClassID<%=Rs("ID")%>">&nbsp;<%=Rs("ClassName")%>��<%
If I Mod 5= 0 Then
Response.Write("<p>")
End if
Rs.MoveNext 
Next
Rs.Close
Set Rs = Nothing
%></td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>

	</div></td>
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