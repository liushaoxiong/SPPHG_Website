<p id="focusViwer" name="focusView">Loading...</p>
<script language="javascript" src="../E_Scripts/su_focusflash.js"></script>
			<script type="text/javascript">			
            var files="<% 
	dim sqlAD,rsAD
	Set rsAD = Server.CreateObject("ADODB.Recordset")
	sqlAD="select top 5 * from E_Ads where AdsContent <> '' And ID > 1 order by Id Asc"

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
%><%=rsAD("AdsContent")%><%=Lqm%>
<%
rsAD.MoveNext 
next
rsAD.Close
set rsAD = nothing
%>";
            var links="<% 
	Set rsAD = Server.CreateObject("ADODB.Recordset")
	sqlAD="select top 5 * from E_Ads where AdsContent <> '' And ID > 1 order by Id Asc"

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
%><%=rsAD("AdsLink")%><%=Lqm%>
<%
rsAD.MoveNext 
next
rsAD.Close
set rsAD = nothing
%>";
            var texts="#|#|#|#|#";
			var ImgHeight="395";
			
			
			var TitleBgColor="0xFFFFFF"	
			var TitleBgAlpha="100"		
			var TitleBgPosition="1000"	
			var BtnDefaultColor="0xFF6600"
			var BtnOverColor="0x0096ff"	
			var AutoPlayTime="5"		
			var Tween="2"
			var	IsShowBtn="1"
			var WinOpen="_blank"
			var FlashVarst= "bcastr_file="+files+"&bcastr_link="+links+"&bcastr_title="+texts+"&TitleBgColor="+TitleBgColor+"&TitleBgAlpha="+TitleBgAlpha+"&TitleBgPosition="+TitleBgPosition+"&BtnDefaultColor="+BtnDefaultColor+"&BtnOverColor="+BtnOverColor+"&AutoPlayTime="+AutoPlayTime+"&Tween="+Tween+"&IsShowBtn="+IsShowBtn+"&WinOpen="+WinOpen
			var FocusFlash = new focusFlash("../E_img/flashpic.swf","focusflash","100%",ImgHeight,"7","#ffffff",false,"High");
			FocusFlash.addParam("allowScriptAccess", "sameDomain");
			FocusFlash.addParam("menu", "false");
			FocusFlash.addParam("wmode", "opaque");
			FocusFlash.addParam("FlashVars", FlashVarst );
			FocusFlash.write("focusViwer");
</script>