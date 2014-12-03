<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
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
    <td width="200" height="600" valign="top" bgcolor="f2f2f2">
    <!--#include file="E_SysMenu.asp"-->
    </td>
    <td width="205" valign="top" background="Img/C_bg.jpg"><!--#include file="SetSiteMenu.html"--></td>
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
              <td class="Wryh T14 P2"><strong>&nbsp;系统信息</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
            <div id="Empty"></div>
            
            <table width="95%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="Wryh T14"><img src="Img/Admin.png" align="absmiddle">　<strong>授权客户：<%=ClientCompany%>　(管理员：<%=session("AUserName")%>)</strong></td>
              </tr>
            </table>
            <div id="Empty"></div>
            <table width="95%" border="0" cellpadding="0" cellspacing="1" bgcolor="#3E4B6B">
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">栏目统计</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_NewsClass")(0)%>&nbsp;条</td>
                <td width="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">新闻统计</td>
                <td width="35%" bgcolor="#FFFFFF">&nbsp;<%=Conn.execute("Select count(Id) From E_News")(0)%>&nbsp;条</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">简介统计</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_About")(0)%>&nbsp;条</td>
                <td width="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">品牌统计</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_Brand")(0)%>&nbsp;个</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">分类统计</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From BClass")(0)%>&nbsp;条</td>
                <td width="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">产品统计</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_product")(0)%>&nbsp;条</td>
              </tr>
              <tr>
                <td width="100" height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">友情链接</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_FriendLink")(0)%>&nbsp;条</td>
                <td width="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">广告管理</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_Ads")(0)%>&nbsp;条</td>
              </tr>
              <tr>
                <td height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">网站管理员</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_Admin")(0)%>&nbsp;位</td>
                <td width="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">前台会员</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_Member")(0)%>&nbsp;位</td>
              </tr>
              <tr>
                <td height="40" align="center" bgcolor="#E3E7EE" class="Wryh T14">网站开通日</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=AuthorizedDate%> | 距今:<%=int(date()-cdate(AuthorizedDate))%>天</td>
                <td width="100" align="center" bgcolor="#E3E7EE" class="Wryh T14">网站留言</td>
                <td bgcolor="#FFFFFF">&nbsp;<%=	Conn.execute("Select count(Id) From E_Message")(0)%>&nbsp;条|<a href="E_MessageMain.asp">未审核&nbsp;<span class="Red"><%=Conn.execute("Select count(Id) From E_Message Where IfShow = 0 ")(0)%></span>&nbsp;条</a></td>
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
