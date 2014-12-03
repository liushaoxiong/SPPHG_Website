<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|8,")=0 then 
  response.write "<script language='JavaScript'>alert('您无权操作此页面，请返回！');" & "history.back()" & "</script>"
  response.end
end if
end if

Default=5
ClassMenu = 2
ClassID = FunIfInt(Request("ClassID"))
ID=FunIfInt(request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from E_Admin where  id="&ID
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
<script language=JavaScript>
function checkdata() {
if (editForm.MagName.value.length < 1) {
alert("\请填写用户名!!")
editForm.MagName.focus()
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
              <td class="Wryh T14 P2"><strong>&nbsp;修改帐号</strong></td>
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
                            <td align="center" class="white Wryh T14"><a href="E_AdminUserMain.asp"><strong>帐号列表</strong></a></td>
                          </tr>
                      </table></td>
                      <td width="10">&nbsp;</td>
                      <td width="103" background="Img/tj.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="30">&nbsp;</td>
                            <td align="center" class="white Wryh T14"><a href="E_AdminUserAdd.asp"><strong>添加用户</strong></a></td>
                          </tr>
                      </table></td>
                      <td>&nbsp;</td>
                    </tr>
                </table></td>
              </tr>
            </table>
            <div id="Empty"></div> 
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=AdminEdit&ID=<%=ID%>" onSubmit="return checkdata()">
              <tr>
                <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">用户名</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagName" class="Titleinput" value="<%=Rs("MagName")%>"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>

              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">密码</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagPass" type="password" class="Titleinput"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">确认密码</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagPass1" type="password" class="Titleinput"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">电子邮箱</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="MagEmail" class="Titleinput" value="<%=Rs("MagEmail")%>"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">管理级别</td>
                <td>
                <select name="MagLevel" class="Input" id="MagLevel" onChange="changeMagLevel(document.editForm.MagLevel.value)">
                <option value="0" <% If Rs("MagLevel") = 0 Then %> selected<%End If%>>超级管理员</option>
                 <option value="1" <% If Rs("MagLevel") = 1 Then %> selected<%End If%>>普通管理员</option></select></td>
              </tr>
              <tr>
                <td width="100" height="50" align="center">&nbsp;</td>
                <td><input type="image" src="Img/Tijiao.jpg"></td>
              </tr>
             
            </table></td>
                <td width="500" valign="top">
			<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"><input name="Purview1" type="checkbox" value="|1," <%if Instr(Rs("MagPopdem"),"|1,")>0 then response.write ("checked")%>>&nbsp;企业信息　<input name="Purview2" type="checkbox" value="|2," <%if Instr(Rs("MagPopdem"),"|2,")>0 then response.write ("checked")%>>&nbsp;编辑企业　<input name="Purview3" type="checkbox" value="|3," <%if Instr(Rs("MagPopdem"),"|3,")>0 then response.write ("checked")%>>&nbsp;新闻管理　
      <p>
        <input name="Purview7" type="checkbox" value="|7," <%if Instr(Rs("MagPopdem"),"|7,")>0 then response.write ("checked")%>>&nbsp;系统管理　<input name="Purview8" type="checkbox" value="|8," <%if Instr(Rs("MagPopdem"),"|8,")>0 then response.write ("checked")%>>&nbsp;帐号管理　<input name="Purview9" type="checkbox" value="|9," <%if Instr(Rs("MagPopdem"),"|9,")>0 then response.write ("checked")%>>&nbsp;静态生成　<input name="Purview10" type="checkbox" value="|10," <%if Instr(Rs("MagPopdem"),"|10,")>0 then response.write ("checked")%>>&nbsp;系统备份<p><input name="Purview11" type="checkbox" value="|11," <%if Instr(Rs("MagPopdem"),"|11,")>0 then response.write ("checked")%>>&nbsp;系统还原　<input name="Purview12" type="checkbox" value="|12," <%if Instr(Rs("MagPopdem"),"|12,")>0 then response.write ("checked")%>>&nbsp;系统备份　<input name="Purview13" type="checkbox" value="|13," <%if Instr(Rs("MagPopdem"),"|13,")>0 then response.write ("checked")%>>&nbsp;广告管理　<input name="Purview14" type="checkbox" value="|14," <%if Instr(Rs("MagPopdem"),"|14,")>0 then response.write ("checked")%>>&nbsp;留言管理　<input name="Purview15" type="checkbox" value="|15," <%if Instr(Rs("MagPopdem"),"|15,")>0 then response.write ("checked")%>>&nbsp;前台会员</td>
  </tr>
    <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#BCFEFE" class="Wryh T14">&nbsp;<strong>新闻分类</strong></td>
  </tr>
  <tr>
    <td height="30"> 
<% 
Set Rs1=Server.Createobject("Adodb.Recordset")
QuerySql="Select * from E_NewsClass order by id Asc"
Rs1.open QuerySql,Conn,1,1
For I = 1 to Rs1.recordcount
if Rs1.EOF then     
Exit For 
end if '利用for next 循环依次读出记录
%><input name="ClassID" type="checkbox" value="|ClassID<%=Rs1("ID")%>" <%if InStr(Rs("MagPopdem"),"ClassID"&Rs1("id")&",")>0 then response.write "checked" end if%>>&nbsp;<%=Rs1("ClassName")%>　<%
If I Mod 5= 0 Then
Response.Write("<p>")
End if
Rs1.MoveNext 
Next
Rs1.Close
Set Rs1 = Nothing
%></td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
</td>
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