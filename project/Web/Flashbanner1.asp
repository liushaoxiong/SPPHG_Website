<p id="focusViwer1" name="focusView1">Loading...</p>
			<script type="text/javascript">			
            var files="<% 
	Set rsAD = Server.CreateObject("ADODB.Recordset")
	sqlAD="select top 5 ID,Title,ClassID,SortID,Pictures from E_News where ClassID = 1 And Pictures <> '' order by SortID Asc,Id Desc"

	rsAD.Open sqlAD, Conn, 1										 
For iRs = 1 to rsAD.Recordcount
if rsAD.EOF then     
Exit For 
end if
If rsAD.Recordcount <> iRs Then
Lqm="|"
Else
Lqm=""
End if
%><%=rsAD("Pictures")%><%=Lqm%>
<%
rsAD.MoveNext 
next
rsAD.Close
set rsAD = nothing
%>";
            var links="<% 
	Set rsAD = Server.CreateObject("ADODB.Recordset")
	sqlAD="select top 5 ID,Title,ClassID,SortID,Pictures from E_News where ClassID = 1 And Pictures <> '' order by SortID Asc,Id Desc"

	rsAD.Open sqlAD, Conn, 1										 
For iRs = 1 to rsAD.Recordcount
if rsAD.EOF then     
Exit For 
end if
If rsAD.Recordcount <> iRs Then
Lqm="|"
Else
Lqm=""
End if
%>/News/E_NewsShow-<%=rsAD("ID")%>-<%=rsAD("ClassID")%>-<%=encode(rsAD("ID"))%>.html<%=Lqm%>
<%
rsAD.MoveNext 
next
rsAD.Close
set rsAD = nothing
%>";
            var texts="#|#|#|#|#";
			var ImgHeight="150";
			
			
			var TitleBgColor="0xFFFFFF"	
			var TitleBgAlpha="100"		
			var TitleBgPosition="1000"	
			var BtnDefaultColor="0xFF6600"
			var BtnOverColor="0x0096ff"	
			var AutoPlayTime="3"		
			var Tween="3"
			var	IsShowBtn="1"
			var WinOpen="_blank"
			var FlashVarst= "bcastr_file="+files+"&bcastr_link="+links+"&bcastr_title="+texts+"&TitleBgColor="+TitleBgColor+"&TitleBgAlpha="+TitleBgAlpha+"&TitleBgPosition="+TitleBgPosition+"&BtnDefaultColor="+BtnDefaultColor+"&BtnOverColor="+BtnOverColor+"&AutoPlayTime="+AutoPlayTime+"&Tween="+Tween+"&IsShowBtn="+IsShowBtn+"&WinOpen="+WinOpen
			var FocusFlash = new focusFlash("../E_img/flashpic.swf","focusflash","180",ImgHeight,"7","#ffffff",false,"High");
			FocusFlash.addParam("allowScriptAccess", "sameDomain");
			FocusFlash.addParam("menu", "false");
			FocusFlash.addParam("wmode", "opaque");
			FocusFlash.addParam("FlashVars", FlashVarst );
			FocusFlash.write("focusViwer1");
</script>