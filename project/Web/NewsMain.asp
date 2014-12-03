<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<%
ClassID =FunIfInt(request("ClassID"))
SearchType =FunIfInt(request("SearchType"))
ScopeID =FunIfInt(request("ScopeID"))
Keywords =Get_SafeStr(request("Keywords"))
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title><%If Keywords <> "" Then %>站内搜索<%Else%><%=FunClassName(ClassID)%><%End if%>|<%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="../Web/Top.html" -->
<table width="1190" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="112" valign="top"><!--#include file="../Web/left.html"--></td>
    <td align="center" valign="top"><table width="96%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="right">您当前的位置：<a href="../">首页</a>&nbsp;>&nbsp;<%If Keywords <> "" Then %>站内搜索<%Else%><%=FunClassName(ClassID)%><%End if%></td>
      </tr>
    </table>
<%Select Case ClassID
Case 9,25
%>
<table width="96%" border="0" cellpadding="0" cellspacing="0" background="../E_img/main_04.jpg">
      <tr>
        <td width="105" valign="top" class="ClassName1" align="center"><%If Keywords <> "" Then %>站内搜索<%Else%><%=FunClassName(ClassID)%><%End if%></td>
        <td height="18"><img src="../E_img/main_04.jpg"></td>
      </tr>
    </table>
<%Case Else%>
    
      <table width="96%" border="0" cellpadding="0" cellspacing="0" background="../E_img/main_04.jpg">
      <tr>
        <td width="85" valign="top" class="ClassName" align="center"><%If Keywords <> "" Then %>站内搜索<%Else%><%=FunClassName(ClassID)%><%End if%></td>
        <td height="18"><img src="../E_img/main_04.jpg"></td>
      </tr>
    </table>
<%End Select%>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20"><%
Set Rs = Server.CreateObject("ADODB.Recordset")
If SearchType = 1 Then
Sql="select * from E_News where Id > 0 And Title Like '%"&Keywords&"%' order by ClassID Asc,ID Desc"
Else
Sql="select * from E_News where ClassID="&ClassID&" order by ClassID Asc,ID Desc"
End If
rs.open sql,conn,1,1
dim CurPage,StartPageNum,EndPageNum,p,sizepage,dispage
filename = "/Web/NewsMain.asp"'此文件的文件名 
Select Case  ClassID
Case 27,29,35,33,36
sizepage =12'设置每页记录数 
Case Else
sizepage =16'设置每页记录数 
End Select
dispage =10'设置页面上显示多少页 
if rs.eof and rs.bof then%>
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
 <%Select Case ClassID%>
      <%Case 27,29,35,33,36%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <%
If EndPageNum>rs.Pagecount then EndPageNum=rs.Pagecount
J=0 
p=RS.PageSize*(Curpage-1) 
do while (Not rs.Eof) and (I<rs.PageSize) 
p=p+1
iRs = I + 1
%>
          <td align="center"><table width="130" border="0" cellspacing="0" cellpadding="0" style="margin:10px">
              <tr>
                <td height="130" align="center" class="Bor"><a href="/Web/NewsShow.asp?ID=<%=Rs("ID")%>" title="<%=Rs("Title")%>" target="_blank"><img src="<%=PictureUrl(Rs("Pictures"))%>" width="120" height="120" /></a></td>
              </tr>
              <tr>
                <td height="35" align="center"><a href="/Web/NewsShow.asp?ID=<%=Rs("ID")%>" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),20)%></a></td>
              </tr>
          </table></td>
          <%
If iRs Mod 5 = 0 Then
Response.Write("</tr><tr><td height=10></td></tr><tr>")
End if
I=I+1 
rs.MoveNext 
Loop
%>
        </tr>
      </table>
  
      <%Case Else%>
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <%
If EndPageNum>rs.Pagecount then EndPageNum=rs.Pagecount
J=0 
p=RS.PageSize*(Curpage-1) 
do while (Not rs.Eof) and (I<rs.PageSize) 
p=p+1
iRs = I + 1
%>
        <tr>
          <td height="30" class="T14"><img src="../E_img/list.jpg" align="absmiddle">&nbsp;&nbsp;<a href="/Web/NewsShow.asp?ID=<%=Rs("ID")%>" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),80)%></a></td>
          <td width="80" align="right" class="Fonten"><%=Year(Rs("AddTime"))%>-<%=right("0"&month(Rs("AddTime")),2)%>-<%=right("0"&day(Rs("AddTime")),2)%></td>
        </tr>
<%
If 	iRs mod 4 = 0 Then
Response.Write("</tr><tr><td height=10></td><td height=10></td></tr><tr><td height=10 class=Dot></td><td height=8 class=Dot></td></tr><tr><td height=8></td><td height=8></td></tr><tr>")
end if	
I=I+1 
rs.MoveNext 
Loop
%>
      </table>
      <%End Select%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table>
      <table width="96%" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="">
          <td height="21" background="../image/0000999.jpg" bgcolor=""><div class="scott">
              <%if CurPage>dispage then%>
              <a href="<%=filename%>?CurPage=1&ClassID=<%=ClassID%>&SearchType=<%=Keywords%>&Keywords=<%=Keywords%>"><font title="首页">&lt;&lt;</font></a>
              <%end if%>
              <%if CurPage>dispage then%>
              <a href="<%=filename%>?CurPage=<%=StartPageNum-1%>&ClassID=<%=ClassID%>&Keywords=<%=Keywords%>&SearchType=<%=Keywords%>"><font title="上<%=dispage%>">&lt;</font></a>
              <%end if 
                  For J=StartPageNum to EndPageNum 
                  if J<>CurPage then %>
              <a href="<%=filename%>?CurPage=<%=J%>&ClassID=<%=ClassID%>&SearchType=<%=Keywords%>&Keywords=<%=Keywords%>"><%=J%></a>
              <% else %>
              <span class="fya"><%=j%></span></font>
              <% end if 
                  Next %>
              <% if EndPageNum<RS.Pagecount then %>
              <a href="<%=filename%>?CurPage=<%=EndPageNum+1%>&ClassID=<%=ClassID%>&SearchType=<%=Keywords%>&Keywords=<%=Keywords%>"><font title="下<%=dispage%>">&gt;</font></a>
              <%end if 
                  if CurPage<TotalPages then%>
              <a href="<%=filename%>?CurPage=<%=TotalPages%>&ClassID=<%=ClassID%>&SearchType=<%=Keywords%>&Keywords=<%=Keywords%>"><font title="尾页">&gt;&gt;</font></a>
              <%end if%>
              <%end if%>
          </div></td>
        </tr>
      </table>
    </td>
    <td width="230" align="right" valign="top">
	<!--#include file="../Web/Right.html"--></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html"-->
</body>
</html>
<% 
Rs.Close
Set Rs = Nothing
%>