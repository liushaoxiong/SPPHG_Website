<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
Default=6
ClassMenu = 1

'========判断是否具有管理权限
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|9,") = 0 then 
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
</head>
<body>
<!--#include file="Top.html"-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" valign="top" bgcolor="f2f2f2">
    <!--#include file="E_SysMenu.asp"-->
    </td>
    <td width="205" valign="top" background="Img/C_bg.jpg">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="Img/jtsc.jpg" width="205" height="40" /></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
        <td <%If ClassMenu = 1 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_HtmlMain.asp"><img src="Img/C2.png" align="absmiddle">　生成列表</a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr>
     
    </table>
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
              <td class="Wryh T14 P2"><strong>&nbsp;静态生成</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
          <div id="Empty"></div>
          <table width="98%" border="0" cellspacing="0" cellpadding="0">
            <tr valign="top">
              <td width="60%" align="right"><table width="98%" border="0" cellspacing="1" cellpadding="0">

                  <tr>
                    <th width="180" height="45" align="left" class="Wryh T14" scope="row">生成html首页</th>
                    <form name="DefaultHtml" action="Admin_htmlindex.asp" method="post">
                      <td height="35" bgcolor="#FAFBF4">&nbsp;
                          <input name="Submit" type="submit" class="Textarea" value="首页生成"></td>
                    </form>
                  </tr>
<tr>
                    <th width="180" height="45" align="left" class="Wryh T14" scope="row">生成简介页面</th>
                    <form name="AboutHtml" action="Admin_htmlAbout.asp" method="post">
                      <td height="35" bgcolor="#FAFBF4">&nbsp;
                          <input name="Submit" type="submit" class="Textarea" value="首页简介页面"></td>
                    </form>
                  </tr>
                  <tr>
                    <th width="180" height="45" align="left" class="Wryh T14" scope="row">批量生成新闻内容的html页</th>
                    <form name="ContentHtml1" action="Admin_htmlnews1.asp" method="post">
                      <td height="35" bgcolor="#FAFBF4">&nbsp;
                          <%call FunListBox("Select * from E_NewsClass where IfShow= 1 order by Id Asc","ClassID","ClassName","Id",ClassID,"请选择栏目","input")%>　<input name="Submit" type="submit" class="Textarea" value="生成">                      </td>
                    </form>
                  </tr>
                  <tr>
                    <th width="180" height="45" align="left" class="Wryh T14" scope="row">全部生成新闻内容html页</th>
                    <form name="ContentallHtml" action="Admin_htmlnews.asp" method="post">
                      <td height="35" bgcolor="#FAFBF4">&nbsp;
                          <input name="Submit" type="submit" class="Textarea" value="生成全部内容页"></td>
                    </form>
                  </tr>
              </table></td>
              <td><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;"><table width="300" border="0" align="center" cellpadding="0" cellspacing="5">
                        <tr>
                          <td width="2" valign="top" bgcolor="#EAEEDB" class="FontGray STYLE2" style="line-height:20px;"></td>
                          <td valign="top" bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;">生成HTML网页的好处是：访问更迅速，更易被搜索引擎抓区到，并排名相对靠前 </td>
                        </tr>
                        <tr>
                          <td width="2" height="1" valign="top"  background="Images/DotLine.gif" bgcolor="#EAEEDB" class="FontGray STYLE2" style="line-height:20px;"></td>
                          <td height="1" valign="top" bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;"  background="Images/DotLine.gif"></td>
                        </tr>
                        <tr>
                          <td width="2" height="1" valign="top"  background="Images/DotLine.gif" bgcolor="#EAEEDB" class="FontGray STYLE2" style="line-height:20px;"></td>
                          <td height="1" valign="top" bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;"  background="Images/DotLine.gif"></td>
                        </tr>
                        <tr>
                          <td width="2" valign="top" bgcolor="#EAEEDB" class="FontGray STYLE2" style="line-height:20px;"></td>
                          <td valign="top" bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;">生成时如果遇到网页正在被使用，或服务器负荷较重时，可能失败，造成网站HTML页面不完整。</td>
                        </tr>
                        <tr>
                          <td width="2" height="1" valign="top"  background="Images/DotLine.gif" bgcolor="#EAEEDB" class="FontGray STYLE2" style="line-height:20px;"></td>
                          <td height="1" valign="top" bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;"  background="Images/DotLine.gif"></td>
                        </tr>
                        <tr>
                          <td width="2" valign="top" bgcolor="#EAEEDB" class="FontGray STYLE2" style="line-height:20px;"></td>
                          <td valign="top" bgcolor="#FAFBF4" class="FontGray" style="line-height:20px;">生成网页的操作将占用大量服务器资源，建议夜间（22:00以后）进行</td>
                        </tr>
                    </table></td>
                  </tr>
              </table></td>
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