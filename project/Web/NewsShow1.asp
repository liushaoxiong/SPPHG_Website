<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<%
ID = int(request("ID"))
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from E_News where id="&id
rsnews.Open sql,conn,1,3
if rsnews.eof and rsnews.bof then
  response.write"<SCRIPT language=JavaScript>alert('找不到此新闻！');"
  response.write"javascript:history.go(-1)</SCRIPT>"
else
 Title=rsnews("Title")
 ClassID=rsnews("ClassID")
 Content=rsnews("Content")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title><%=Title%>|<%=FunClassName(ClassID)%>|<%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
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
            <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle"><%If Keywords <> "" Then %>
              站内搜索
                    <%Else%>
              <%=FunClassName(ClassID)%>
              <%End if%></td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="245" align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="50" align="center" class="T20 Dot"><strong><%=Title%></strong></td>
          </tr>
          <tr>
            <td height="40" align="left">&nbsp;浏览次数：
              <script src="/web/VisitClick.asp?id=<%=id%>"></script>
              次 | 发布日期:<%=rsnews("AddTime")%>&nbsp;&nbsp;<a href="/Web/NewsMain.asp?ClassID=<%=ClassID%>">&nbsp;返回上级</a>&nbsp; </td>
          </tr>
          <tr>
            <td height="20" align="left"></td>
          </tr>
          <tr>
            <td align="left" class="T14 line25"><%=Content%></td>
          </tr>
          <tr>
            <td class="line25 Dot">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" align="right" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="right"><%
set RelationArt = conn.execute("select top 1 ID from E_News where ClassID="&ClassID&" And id<"&id&"  order by Id desc")
IF RelationArt.eof and relationArt.bof Then
Response.Write "没有了"    
else
Response.Write "<a href=""/News/E_NewsShow-" & RelationArt(0) & "-" & ClassID & "-" & encode(RelationArt(0)) & ".html"">上一条 PREV</a>"   
end if   
Set RelationArt=Nothing
%>
                    |
                    <%set RelationArt = conn.execute("select top 1 ID from E_News where ClassID="&ClassID&" And  id>"&id&" order by id")
IF RelationArt.eof and relationArt.bof Then
Response.Write "没有了"   
else
Response.Write "<a href=""/News/E_NewsShow-" & RelationArt(0) & "-" & ClassID & "-" & encode(RelationArt(0)) & ".html"">下一条 NEXT</a>"  
end if  
RelationArt.close
Set RelationArt=Nothing 
%>
                    &nbsp; </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="50">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30">&nbsp;</td>
          </tr>
      </table></td>
    <td width="8">&nbsp;</td>
  </tr>
</table>
<!--#include file="../Web/Foot.html" -->
</body>
</html>
<%
  end if
  rsnews.close
  set rsnews=nothing
%>