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
<table width="1190" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="112" valign="top"><!--#include file="../Web/left.html"--></td>
    <td align="center" valign="top"><table width="96%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="right">您当前的位置：<a href="../">首页</a>&nbsp;>&nbsp;
              <%If Keywords <> "" Then %>
          站内搜索
          <%Else%>
          <%=FunClassName(ClassID)%>
          <%End if%></td>
      </tr>
    </table>
        <%Select Case ClassID
Case 9,25
%>
        <table width="96%" border="0" cellpadding="0" cellspacing="0" background="../E_img/main_04.jpg">
          <tr>
            <td width="105" valign="top" class="ClassName1" align="center"><%If Keywords <> "" Then %>
              站内搜索
                <%Else%>
              <%=FunClassName(ClassID)%>
              <%End if%></td>
            <td height="18"><img src="../E_img/main_04.jpg"></td>
          </tr>
        </table>
      <%Case Else%>
        <table width="96%" border="0" cellpadding="0" cellspacing="0" background="../E_img/main_04.jpg">
          <tr>
            <td width="85" valign="top" class="ClassName" align="center"><%If Keywords <> "" Then %>
              站内搜索
                <%Else%>
              <%=FunClassName(ClassID)%>
              <%End if%></td>
            <td height="18"><img src="../E_img/main_04.jpg"></td>
          </tr>
        </table>
      <%End Select%>
        
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="50" align="center" class="T20 Dot"><strong><%=Title%></strong></td>
          </tr>
          <tr>
            <td height="40" align="center">&nbsp;浏览次数：
              <script src="/web/VisitClick.asp?id=<%=id%>"></script>
              次 | 发布日期:<%=rsnews("AddTime")%>&nbsp;&nbsp;<a href="/Web/NewsMain.asp?ClassID=<%=ClassID%>">&nbsp;返回上级</a>&nbsp; </td>
          </tr>
          <tr>
            <td height="20" align="left"></td>
          </tr>
          <%If ClassID = 33 Then%>
          <tr>
            <td align="center"><embed src="http://static.youku.com/v/swf/qplayer.swf?winType=adshow&VideoIDS=<%=rsnews("Title1")%>&isAutoPlay=true&ShowRelatedVideo=True" width="680″ height=" height="460" type="application/x-shockwave-flash" id="movie_player" name="movie_player" bgcolor="#FFFFFF" quality="high" wmode="transparent" allowfullscreen="True"
flashvars="isAutoPlay=True&ShowRelatedVideo=True"
pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></td>
          </tr>
          <tr>
            <td height="30"></td>
          </tr>
          <%End if%>
          <tr>
            <td align="left" class="T14 line25"><%=Content%></td>
          </tr>
          <tr>
            <td class="line25 Dot">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" align="right" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                  <td><%
set RelationArt = conn.execute("select top 1 ID from E_News where ClassID="&ClassID&" And id<"&id&"  order by Id desc")
IF RelationArt.eof and relationArt.bof Then
Response.Write "没有了"    
else
Response.Write "<a href=""/Web/NewsShow.asp?id="& RelationArt(0) & """>上一条</a>"   
end if   
Set RelationArt=Nothing
%>
                    |
                    <%set RelationArt = conn.execute("select top 1 ID from E_News where ClassID="&ClassID&" And  id>"&id&" order by id")
IF RelationArt.eof and relationArt.bof Then
Response.Write "没有了"   
else
Response.Write "<a href=""/Web/NewsShow.asp?id="& RelationArt(0) & """>下一条</a>"  
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
    <td width="230" align="right" valign="top"><!--#include file="../Web/Right.html"--></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html"-->
</body>
</html>
<%
  end if
  rsnews.close
  set rsnews=nothing
%>