<table width="100%" border="0" cellpadding="0" cellspacing="0" background="Img/Search.jpg">
 <Form name="Form" method="post" action="E_NewsMain.asp?ClassID=<%=ClassID%>&MType=1">   
      <tr>
      <td width="18" height="61">&nbsp;</td>
      <td height="61"><input name="Key" style="width:100%; background:none; border:none; color:#fff; padding-top:2px" onFocus="if(this.value==' 新闻搜索!') {this.value='';}this.style.color='#fff';" onBlur="if(this.value=='') {this.value=' 新闻搜索!';this.style.color='#fff';}" value=" 新闻搜索!"   onKeyUp="keyupdeal(event);" onKeyDown="keydowndeal(event);" onClick="keyupdeal(event);"></td>
        <td width="40"><input type="image" src="Img/Ap.gif"></td>
    </tr>
  </Form>      
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><a href="E_NewsMain.asp"><img src="Img/xwlm.jpg" width="205" height="40"></a></td>
      </tr>
<%
Set RsMenu = Server.CreateObject("ADODB.Recordset")
	SqlRs="select  * from E_NewsClass where  IfShow = 1 order by Id Asc"
RsMenu.open SqlRs,conn,1,1 
For iRs = 1 to RsMenu.Recordcount
if RsMenu.EOF then     
Exit For 
end if
%>     
      <tr>
        <td <%If (ClassID - RsMenu("ID")) = 0 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_NewsMain.asp?ClassID=<%=RsMenu("ID")%>"><img src="Img/C3.png" align="absmiddle">　<%=RsMenu("ClassName")%></a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr>
<%
RsMenu.MoveNext 
next
RsMenu.Close
Set RsMenu = nothing
%>      
    </table>