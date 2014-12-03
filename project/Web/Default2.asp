<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<%Default=1%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title><%=ClientkeyWords%>-<%=ClientTitle%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../E_Scripts/Faq.js"></script>
</head>
<body>
<!--#include file="../Web/Top.html"-->
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="Warp">
  <tr>
    <td height="20"></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="Warp">
  <tr>
    <td width="8"></td>
    <td height="240" valign="top"><table width="490" border="0" cellspacing="0" cellpadding="0" class="fenseBro">
      <tr>
        <td><img src="../E_img/qm_22.jpg" width="239" height="34"></td>
        <td align="right"><a href="/Web/NewsMain.asp?Classid=1"><img src="../E_img/qm_24.jpg" width="48" height="34"></a></td>
      </tr>
    </table>
      <table width="490" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="13"></td>
        </tr>
      </table>
      <table width="490" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="210"><table width="190" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="160" align="center" bgcolor="f0d1ee"><!--#include file="../Web/Flashbanner1.asp"--></td>
            </tr>
          </table></td>
          <td valign="top"><table width="95%" border="0" cellspacing="0" cellpadding="0">
            <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 6 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 1 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
            <tr>
              <td height="27"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),30)%></a></td>
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
    <td width="450" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="55" background="../E_img/qm_26.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="140">&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td class="white T24"><%=ClientMobile%></td>
                  </tr>
                  <tr>
                    <td class="white T16"><%=ClientTelephone%></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
            <td width="95" height="55"><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=ClientOicq%>&site=qq&menu=yes"><img src="../E_img/ap.gif" width="88" height="47"></a></td>
          </tr>
        </table></td>
      </tr>
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="15"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="13"><a href="/About/About-1.Html"><img src="../E_img/qm_32.jpg" width="453" height="23"></a></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="13"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="10" cellspacing="0" bgcolor="f2f2f2">
        <tr>
          <td height="100" class="Line22" valign="top">
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 1 * from E_About where ID = 1"
Rs.open SqlRs,conn,1,1 
%>
<%=StrLeft(Rs("TiTle1"),264)%>
<% 
Rs.Close
Set Rs  = Nothing
%>
          </td>
        </tr>
      </table></td>
  </tr>
</table>

<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="Warp">
  <tr>
    <td height="410" align="center" valign="top" bgcolor="fff9ff"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
      <table width="98%" border="0" cellpadding="0" cellspacing="0" background="../E_img/qm_36.jpg">
        <tr>
          <td><img src="../E_img/qm_35.jpg" width="229" height="45"></td>
          <td align="right" valign="bottom"><a href="/Web/NewsMain.asp?Classid=2&ScopeID=1"><img src="../E_img/almore.jpg" width="42" height="27"></a></td>
        </tr>
      </table>
      <table width="98%" border="0" cellpadding="0" cellspacing="0" class="fenseBro1">
        <tr>
          <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 4 ID,Title,Title1,ClassID,SortID,AddTime,ScopeID,Pictures from E_News where ClassID = 2 And ScopeID = 1 And Pictures <> '' order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
          <td height="330" align="center"><table width="210" border="0" cellpadding="0" cellspacing="0" class="huiseBro">
            <tr>
              <td height="240"><a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><img src="<%=Rs("Pictures")%>" width="210" height="240"></a></td>
            </tr>
            <tr>
              <td height="45"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50" valign="bottom" bgcolor="19161c" class="white">NO.<span class="T26"><%=iRs%></span></td>
                  <td height="45" bgcolor="565656" class="white T14">&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),14)%></a></td>
                </tr>
              </table></td>
            </tr>
          </table></td>
<% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
        </tr>
      </table></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="Warp">
  <tr>
    <td height="20"></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="Warp">
  <tr>
    <td width="8"></td>
    <td width="365" height="360" valign="top"><table width="356" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><a href="/About/About-2.Html"><img src="../E_img/qm_43.jpg" width="356" height="44"></a></td>
      </tr>
    </table>
      <table width="356" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../E_img/qm_51.jpg" width="356" height="132" border="0" usemap="#Map"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="22">&nbsp;</td>
        </tr>
      </table>
      <table width="340" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="106" align="right" valign="top" background="../E_img/qm_54.jpg"><table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="30" valign="bottom"><a href="/About/About-2.html"><span class="zise T14">加盟条件</span></a></td>
            </tr>
            <tr>
              <td height="30" valign="bottom"><a href="/About/About-3.html"><span class="zise T14">加盟流程</span></a></td>
            </tr>
            <tr>
              <td height="30" valign="bottom"><a href="/About/About-4.html"><span class="zise T14">加盟展示</span></a></td>
            </tr>
          </table></td>
        </tr>
      </table></td>
      <td width="1" bgcolor="#e7e2e6"></td>
    <td align="center" valign="top"><table width="289" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><a href="/Web/MessageList.asp"><img src="../E_img/qm_45.jpg" width="289" height="42"></a></td>
      </tr>
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="10"></td>
        </tr>
      </table>
      <table width="290" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"></td>
        </tr>
      </table>
      <table width="290" border="0" cellspacing="0" cellpadding="0">
        <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 5  * from E_Message where IfShow = 1 And ReContent <> ''  order by Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
        <tr>
          <td height="33" <%If iRs Mod 2 <> 0 Then%> bgcolor="#ffe0fd"<%End if%>>&nbsp;<span class="T14 zise"><%=iRs%>.</span>&nbsp;<a onClick=javascript:ShowFLT(<%=iRs%>) href="javascript:void(null)"><span class="MessTitle T14"><%=StrLeft(Rs("MesTitle"),34)%></span></a></td>
        </tr>
        <tr id=LM<%=iRs%> <%If iRs = 1 Then%><%Else%> style="DISPLAY: none"<%End if%>>
          <td height="80" class="MessBg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="MessBg1" width="10">答：</td>
    <td class="MessBg1"><%=StrLeft(Rs("ReContent"),110)%></td>
  </tr>
</table>
</td>
        </tr>
<% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
      </table></td>
    <td width="1" bgcolor="#e7e2e6"></td>
    <td width="260" align="center" valign="top"><table width="242" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><a href="/Web/NewsMain.asp?Classid=3"><img src="../E_img/qm_47.jpg" width="242" height="42"></a></td>
      </tr>
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="14"></td>
        </tr>
      </table>
      <table width="242" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="130" align="center" bgcolor="#333333"><%Call Advertisement(1,230,120)%></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="14"></td>
        </tr>
      </table>
      <table width="93%" border="0" cellspacing="0" cellpadding="0">
        <%
Set Rs = Server.CreateObject("ADODB.Recordset")
	SqlRs="select top 4 ID,Title,Title1,ClassID,SortID,AddTime from E_News where ClassID = 3 order by SortID Asc,Id Desc"
Rs.open SqlRs,conn,1,1 
For iRs = 1 to Rs.Recordcount
if Rs.EOF then     
Exit For 
end if
%>
        <tr>
          <td height="27" class="T14"><strong class="T16">&middot;</strong>&nbsp;&nbsp;<a href="/News/E_NewsShow-<%=Rs("ID")%>-<%=Rs("ClassID")%>-<%=encode(Rs("ID"))%>.html" title="<%=Rs("Title")%>" target="_blank"><%=StrLeft(Rs("Title"),27)%></a></td>
        </tr>
        <% 
Rs.movenext
next
Rs.Close
Set Rs = Nothing
%>
      </table></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html"-->
<map name="Map">
<area shape="circle" coords="63,72,51" href="/About/About-3.Html"><area shape="circle" coords="179,75,49" href="/About/About-2.Html">
<area shape="circle" coords="289,73,48" href="/About/About-4.Html">
</map></body>
</html>