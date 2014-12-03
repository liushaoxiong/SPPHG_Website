<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<!--#include file="../E_Inc/E_BClassSClass.asp"-->
<%
Default=3
ClassMenu = 2

'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|4,")=0 then 
  response.write "<script language='JavaScript'>alert('您无权操作此页面，请返回！');" & "history.back()" & "</script>"
  response.end
end if
end if


ID = FunIfInt(Request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from E_product where  id="&ID
Rs.Open sql,conn,1,1
if Rs.eof and Rs.bof then
response.Write("没有记录")
else
BClass = Rs("BClass")
SClass = Rs("SClass")
SortID = Rs("SortID")
Brand = Rs("Brand")
d_savefilename = Rs("d_savefilename")
d_originalfilename = Rs("d_originalfilename")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网络管理后台</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="E_Admin.js"></script>
<script language=JavaScript>
function checkdata() {
if (editForm.ProName.value.length < 1) {
alert("\请填写产品名称!!")
editForm.ProName.focus()
return false;
}
if (editForm.BClass.value == '0') {
alert("\产品分类不能为空，请返回选择!!")
editForm.BClass.focus()
return false;
}
if (editForm.Pictures.value == '') {
alert("\产品图片不能为空，请返回选择!!")
editForm.Pictures.focus()
return false;
}
if (eWebEditor1.getHTML()==""){
      alert("内容不能为空，请返回输入！");
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
              <td class="Wryh T14 P2"><strong>&nbsp;产品修改</strong></td>
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
                        <td align="center" class="white Wryh T14"><a href="E_ProductMain.asp"><strong>产品列表</strong></a></td>
                      </tr>
                    </table></td>
                    <td width="10">&nbsp;</td>
                    <td width="103" background="Img/tj.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="30">&nbsp;</td>
                        <td align="center" class="white Wryh T14"><a href="E_ProductAdd.asp"><strong>添加产品</strong></a></td>
                      </tr>
                    </table></td>
                     <td width="10">&nbsp;</td>
                    <td width="103" background="Img/Back.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="30">&nbsp;</td>
                        <td align="center" class="white Wryh T14"><a href="Javascript:history.back();"><strong>返回上级</strong></a></td>
                      </tr>
                    </table></td> 
                    <td>&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
            </table>
             <div id="Empty"></div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=ProEdit&ID=<%=ID%>" onSubmit="return checkdata()">
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">产品名称</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="ProName" class="Titleinput" value="<%=Rs("ProName")%>"></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">产品分类</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="240"><%Call BClass2SClass(BClass,SClass,"editForm","input",1)%></td>
    <td class="Wryh T14">所属品牌：<%call FunListBox("Select * from E_Brand  order by Id Asc","Brand","BrandName","Id",Brand,"请选择品牌","input")%></td>
  </tr>
</table>
</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">市场价</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="80"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="Price1" class="Titleinput" <%=rs("Price1")%>></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
                <td width="80" align="center" class="Wryh T14">会员价</td>
    <td class="Wryh T14" width="80"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="Price2" class="Titleinput" <%=rs("Price2")%>></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
                <td>&nbsp;</td>
  </tr>
</table>
</td>
              </tr>
              <tr>
                <td width="100" height="100" align="center" class="Wryh T14">产品摘要</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/input1_01.jpg" width="8" height="86"></td>
                    <td background="Img/Input1_03.jpg"><textarea name="Title" rows="4" class="Title1input"><%If Rs("Title") <> "" Then%><%=changechr2(Rs("Title"))%><%Else%><%=Rs("Title")%><%End if%></textarea></td>
                    <td width="8"><img src="Img/input1_05.jpg" width="8" height="86"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">产品小图</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                    <td background="Img/Input_03.jpg"><input name="Pictures" class="Titleinput" <%=rs("Pictures")%>></td>
                    <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                  </tr>
                </table></td>
    <td width="20%"><a href="javaScript:OpenScript('UpFileForm.asp?Result=Pictures',460,180)"><img src="Img/Upload.jpg" width="91" height="30" border="0" align="absmiddle"></a></td>
    <td width="10%">&nbsp;</td>
  </tr>
</table>
</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">产品排序</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="60"><select name="SortID" class="Input">
             <option value="9999"<%if SortID =9999 then response.write "selected" end if%>>默认</option>
              <%for i =1 to 30%>
              <option value="<%=i%>" <%if SortID = i then response.Write("Selected") end if%>><%=i%></option>
              <%next%>
            </select></td>
    <td class="Wryh T14" width="70">新品：
      <input name="NewFlag" type="checkbox" value="1"></td>
    <td class="Wryh T14" width="100">推荐：<input name="Recommend" type="checkbox" value="1"></td>
    <td>&nbsp;</td>
  </tr>
</table>
</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" class="Wryh T14">产品内容</td>
                <td><textarea name="Content" id="Content" style="display:none"><%=Server.HTMLEncode((Rs("Content")))%></textarea>
        <iframe ID="eWebEditor1" src="../E_eWebEditor/ewebeditor.htm?id=Content&style=Mini&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename" frameborder="0" scrolling="no" width="90%" height="350"></iframe></td>
              </tr>
              <tr>
                <td width="80" height="50" align="center">&nbsp;</td>
                <td>
                <input type=hidden name=d_originalfilename id=d_originalfilename>
                <input type=hidden name=d_savefilename id=d_savefilename>
                <input type="image" src="Img/Tijiao.jpg"></td>
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