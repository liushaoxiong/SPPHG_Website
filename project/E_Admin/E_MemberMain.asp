<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
Default=4
ClassMenu = 4

'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|15,")=0 then 
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
<script language="javascript">
function ifmsgbox(message,url) {
	     if ( confirm(message) ) {
		 window.location.href=url
		 }
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
              <td class="Wryh T14 P2"><strong>&nbsp;帐号管理</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" background="Img/Add_banner.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="20">&nbsp;</td>
                      <td width="103" height="27" background="Img/qb.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="30">&nbsp;</td>
                            <td align="center" class="white Wryh T14"><a href="E_AdminUserMain.asp"><strong>帐号列表</strong></a></td>
                          </tr>
                      </table></td>
                      <td width="10">&nbsp;</td>
                      <td width="103" background="Img/tj1.jpg" style="display:none"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
            <table width="100%" border="0" cellpadding="0" cellspacing="0" background="Img/Banner.jpg">
              <tr>
                <td width="50" align="center" class="Wryh T13">ID</td>
                <td width="100" height="40" align="center" class="Wryh T13">用户名</td>
                <td align="right" class="Wryh T13">最近登陆日期</td>
                <td width="100" align="right" class="Wryh T13">最近登陆IP</td>
                <td width="100" align="center" class="Wryh T13">操作</td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><%
Set Rs = Server.CreateObject("ADODB.Recordset")
Sql="select * from E_Member order by ID Asc"
rs.open sql,conn,1,1
dim CurPage,StartPageNum,EndPageNum,p,sizepage,dispage
filename = "E_MemberMain.asp"'此文件的文件名 
sizepage =10'设置每页记录数 
dispage =10'设置页面上显示多少页 

if rs.eof and rs.bof then%>
<br>
                  　
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
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <%
If EndPageNum>rs.Pagecount then EndPageNum=rs.Pagecount
I=0 
p=RS.PageSize*(Curpage-1) 
do while (Not rs.Eof) and (I<rs.PageSize) 
p=p+1
iRs = I + 1
%>
              <tr onMouseOut="this.style.backgroundColor='#ffffff'" onMouseOver="this.style.backgroundColor='<%If iRs mod 2 = 0 then response.Write("#d5ffff") Else response.Write("#f8ffd1") End if%>'" style="cursor:hand">
                <td width="50" height="35" align="center" ><strong><%=iRs%></strong></td>
                <td width="100" align="center"><strong class="Blue"><%=Rs("UserName")%></strong></td>
                <td align="right" class="Time"><%=Rs("lstTime")%></td>
                <td align="right" width="100"><%=Rs("lstIp")%></td>
                <td width="100" align="center"><a href="E_AdminUserEdit.asp?ID=<%=Rs("ID")%>"><img src="Img/Edit.Gif" width="16" height="16"  alt="修改"></a>&nbsp;&nbsp;<a href="#" onClick="ifmsgbox('将本条内容彻底删除吗？\n\n删除不可恢复，其相关的文件也将一并删除！','E_ManageSave.asp?SaveType=AdminDel&Id=<%=Rs("Id")%>&CurPage=<%=CurPage%>');"><img src="Img/Del.gif"  alt="删除"></a></td>
              </tr>
                <td height="1" class="Dot"></td>
                <td height="1" class="Dot"></td>
                <td height="1" class="Dot"></td>
                <td height="1" class="Dot"></td>
                <td height="1" class="Dot"></td>
              </tr>
              <%
I=I+1 
rs.MoveNext 
Loop
%>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr bgcolor="">
                <td height="21"><div class="scott">分页：
                  <%if CurPage>dispage then%>
                        <a href="<%=filename%>?CurPage=1"><font title="首页"><<</font></a>
                        <%end if%>
            <%if CurPage>dispage then%>
                        <a href="<%=filename%>?CurPage=<%=StartPageNum-1%>"><font title="上<%=dispage%>"><</font></a>
                        <%end if 
                  For I=StartPageNum to EndPageNum 
                  if I<>CurPage then %>
                        <a href="<%=filename%>?CurPage=<%=I%>"><%=I%></a>
            <% else %>
                        <span class="fya"><%=I%></span></font>
                        <% end if 
                  Next %>
            <% if EndPageNum<RS.Pagecount then %>
                        <a href="<%=filename%>?CurPage=<%=EndPageNum+1%>"><font title="下<%=dispage%>">></font></a>
                        <%end if 
                  if CurPage<TotalPages then%>
                        <a href="<%=filename%>?CurPage=<%=TotalPages%>"><font title="尾页">>></font></a>
                        <%end if%>
                        <%end if%>
                </div></td>
                <td width="20">&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
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
Rs.Close
Set Rs = Nothing
%>