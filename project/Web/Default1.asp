<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title><%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="../Web/Top.html"-->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="415" align="center" valign="top"><table width="97%" border="0" cellpadding="0" cellspacing="0" class="fffeceBor">
      <tr>
        <td height="270" align="center"><!--#include file="../Web/Flashbanner.asp"--></td>
      </tr>
    </table></td>
    <td valign="top" style="padding-left:5px"><table width="97%" border="0" cellpadding="0" cellspacing="0" class="fffeceBor">
      <tr>
        <td height="270" align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="10"></td>
          </tr>
        </table>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 1 ID,Title,Title1,ClassID,SortID from E_News where ClassID = 2 And Title1 <> '' order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
          <tr>
            <td height="30"><a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><strong><%=StrLeft(Rs("Title"),42)%></strong></a></td>
          </tr>
          <tr>
            <td height="80" valign="top" class="Bottomline Line22"><%=StrLeft(Rs("Title1"),142)%>　<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><span class="Red">[了解详情]</span></a></td>
          </tr>
<% 
Rs.movenext
next
Rs.Close
Set Rs  = Nothing
%>
        </table>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="10"></td>
            </tr>
          </table>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 5 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 2 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
            <tr>
              <td height="25"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),44)%></a></td>
              <td width="40" align="right" class="Time"><%=right("0"&month(Rs("AddTime")),2)%>-<%=right("0"&day(Rs("AddTime")),2)%></td>
            </tr>
            <% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
          </table></td>
      </tr>
    </table></td>
    <td width="215" valign="top"><table width="98%" border="0" cellpadding="0" cellspacing="0" class="Bor">
      <tr>
        <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_25.jpg">
          <tr>
            <td width="41"><img src="../E_img/ax_24.jpg" width="41" height="34"></td>
            <td class="RedTitle">关于协会</td>
            <td width="50"><a href="/About/About-1.html"><span class="Red">更多&gt;&gt;</span></a></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="235" align="center" valign="top"><table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="10"></td>
          </tr>
        </table>
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td class="Line22">
<div id="demoqq" style="OVERFLOW: hidden; WIDTH:100%; HEIGHT:215px;">
<div id="demoqq1"> 
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
SqlRs="select ID,TiTle1 from E_About where ID = 1"
Rs.open SqlRs,conn,1,1 
%>
<a href="/About/About-<%=Rs("ID")%>.html"><%=Rs("TiTle1")%></a>
<% 
Rs.Close
Set Rs = Nothing
%>
</div>
<div id="demoqq2"></div>
</div>
<script language="JavaScript" type="text/javascript"> 
var speed=50 
demoqq2.innerHTML=demoqq1.innerHTML 
function Marqueeuq(){ 
if(demoqq2.offsetTop-demoqq.scrollTop<=0) 
demoqq.scrollTop-=demoqq1.offsetHeight 
else{ 
demoqq.scrollTop++ 
} 
} 
var MyMaruq=setInterval(Marqueeuq,speed) 
demoqq.onmouseover=function() {clearInterval(MyMaruq)} 
demoqq.onmouseout=function() {MyMaruq=setInterval(Marqueeuq,speed)} 
</script></td>
            </tr>
          </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="150" align="center"><%Call Advertisement(1,990,125)%></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="8" valign="top">&nbsp;</td>
    <td width="225" valign="top"><!--#include file="../Web/left.html"--></td>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" valign="top"><table width="98%" border="0" cellpadding="0" cellspacing="0" class="Bor">
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_35.jpg">
              <tr>
                <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle">项目资助</td>
                <td>&nbsp;</td>
                <td width="50"><a href="/Web/NewsMain.asp?ClassID=7"><span class="Red">更多&gt;&gt;</span></a></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="245" align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="10"></td>
              </tr>
            </table>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
              <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 9 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 7 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
              <tr>
                <td height="25"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),44)%></a></td>
                <td width="40" align="right" class="Time"><%=right("0"&month(Rs("AddTime")),2)%>-<%=right("0"&day(Rs("AddTime")),2)%></td>
              </tr>
              <% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
            </table></td>
          </tr>
        </table></td>
        <td width="50%" align="right" valign="top"><table width="98%" border="0" cellpadding="0" cellspacing="0" class="Bor">
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_35.jpg">
                <tr>
                  <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle">结对资助</td>
                  <td>&nbsp;</td>
                  <td width="50"><a href="/Web/NewsMain.asp?ClassID=8"><span class="Red">更多&gt;&gt;</span></a></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="245" align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="10"></td>
                </tr>
              </table>
                <table width="95%" border="0" cellspacing="0" cellpadding="0">
                  <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 9 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 8 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
                  <tr>
                    <td height="25"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),44)%></a></td>
                    <td width="40" align="right" class="Time"><%=right("0"&month(Rs("AddTime")),2)%>-<%=right("0"&day(Rs("AddTime")),2)%></td>
                  </tr>
                  <% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
              </table></td>
          </tr>
        </table></td>
      </tr>
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="15">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" valign="top"><table width="98%" border="0" cellpadding="0" cellspacing="0" class="Bor">
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_35.jpg">
                    <tr>
                      <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle">专题活动</td>
                      <td>&nbsp;</td>
                      <td width="50"><a href="/Web/NewsMain.asp?ClassID=5"><span class="Red">更多&gt;&gt;</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="245" align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="10"></td>
                    </tr>
                  </table>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 9 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 5 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
                      <tr>
                        <td height="25"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),44)%></a></td>
                        <td width="40" align="right" class="Time"><%=right("0"&month(Rs("AddTime")),2)%>-<%=right("0"&day(Rs("AddTime")),2)%></td>
                      </tr>
                      <% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
                  </table></td>
              </tr>
          </table></td>
          <td width="50%" align="right" valign="top"><table width="98%" border="0" cellpadding="0" cellspacing="0" class="Bor">
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../E_img/ax_35.jpg">
                    <tr>
                      <td width="131" height="37" align="center" background="../E_img/ax_34.jpg" class="RedTitle">本地新闻</td>
                      <td>&nbsp;</td>
                      <td width="50"><a href="/Web/NewsMain.asp?ClassID=6"><span class="Red">更多&gt;&gt;</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="245" align="center" valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="10"></td>
                    </tr>
                  </table>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 9 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 6 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
                      <tr>
                        <td height="25"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),44)%></a></td>
                        <td width="40" align="right" class="Time"><%=right("0"&month(Rs("AddTime")),2)%>-<%=right("0"&day(Rs("AddTime")),2)%></td>
                      </tr>
                      <% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
                  </table></td>
              </tr>
          </table></td>
        </tr>
      </table></td>
    <td width="8" valign="top">&nbsp;</td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="20"></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="8"></td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td background="../E_img/ax_79.jpg"><img src="../E_img/ax_69.jpg" width="200" height="30"></td>
        <td width="51"><a href="/Web/NewsMain.asp?ClassID=4"><img src="../E_img/ax_73.jpg" width="51" height="30"></a></td>
      </tr>
    </table></td>
    <td width="8"></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="8"></td>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="UnBottom">
      <tr>
        <td height="190" align="center">    <DIV id=ddemoq style="OVERFLOW: hidden; WIDTH: 98%;">
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0 cellspace="0">
        <TBODY>
        <TR>
          <TD id=ddemoq1 vAlign=top align="center">
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr>
                      <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 15 ID,Title,Title1,ClassID,SortID,AddTime,Pictures from E_News where ClassID = 4 And Pictures <> '' order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
    <td ><table width="220" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="205" height="140" align="center" class="Bor">
   <a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><img src="<%=Rs("Pictures")%>" width="195" height="130"></a></td>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td height="30" align="center"><a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),26)%></a></td>
    <td id="jia">&nbsp;</td>
  </tr>
  
</table>
</td>
<% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
  </tr>
</table>

          </TD>
<TD id=ddemoq2>&nbsp;</TD></TR></TBODY></TABLE></DIV>
<script language="javascript">
var speed3=25//速度数值越大速度越慢
ddemoq2.innerHTML=ddemoq1.innerHTML
function dMarqueeq(){
if(ddemoq2.offsetWidth-ddemoq.scrollLeft<=0)
ddemoq.scrollLeft-=ddemoq1.offsetWidth
else{
ddemoq.scrollLeft++
}
}
var dMyMarq=setInterval(dMarqueeq,speed3)
ddemoq.onmouseover=function() {clearInterval(dMyMarq)}
ddemoq.onmouseout=function() {dMyMarq=setInterval(dMarqueeq,speed3)}
</script></td>
      </tr>
    </table></td>
    <td width="8"></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html"-->
</body>
</html>