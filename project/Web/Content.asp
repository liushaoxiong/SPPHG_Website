		<%Select Case SortID
		Case 6,7,8,9,10,11,12,13,14,15,16,17,18,19
		%>
        <%TName="开业庆典"%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    <td bgcolor="fcc8e6" class="Fonten Line22" style="padding:10px">
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from Southidc_About where id=13"
rsnews.Open sql,conn,1,3
%>
<%=rsnews("Content")%>
<%
  rsnews.close
  set rsnews=nothing
%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
        <%Case 20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39%>
        <%TName="开工奠基"%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    <td bgcolor="fcc8e6" class="Fonten Line22" style="padding:10px">
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from Southidc_About where id=14"
rsnews.Open sql,conn,1,3
%>
<%=rsnews("Content")%>
<%
  rsnews.close
  set rsnews=nothing
%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
        <%Case 40,41,42,43%>
        <%TName="礼仪模特"%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    <td bgcolor="fcc8e6" class="Fonten Line22" style="padding:10px">
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from Southidc_About where id=15"
rsnews.Open sql,conn,1,3
%>
<%=rsnews("Content")%>
<%
  rsnews.close
  set rsnews=nothing
%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
        <%Case 44,45,46,47,48,49,50%>
        <%TName="婚礼服务"%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    <td bgcolor="fcc8e6" class="Fonten Line22" style="padding:10px">
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from Southidc_About where id=16"
rsnews.Open sql,conn,1,3
%>
<%=rsnews("Content")%>
<%
  rsnews.close
  set rsnews=nothing
%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
        <%Case 51,52,53,54,55,56%>
        <%TName="会议策划"%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    <td bgcolor="fcc8e6" class="Fonten Line22" style="padding:10px">
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from Southidc_About where id=17"
rsnews.Open sql,conn,1,3
%>
<%=rsnews("Content")%>
<%
  rsnews.close
  set rsnews=nothing
%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
        <%Case 2,3,4%>
        <%TName="晚会演出"%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    <td bgcolor="fcc8e6" class="Fonten Line22" style="padding:10px">
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet")
sql="select * from Southidc_About where id=13"
rsnews.Open sql,conn,1,3
%>
<%=rsnews("Content")%>
<%
  rsnews.close
  set rsnews=nothing
%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
        <%End Select%>