<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<%
SaveType = Get_SafeStr(Trim(Request("SaveType")))

Select Case  SaveType

'�����ҵ��Ϣ
Case "MessageAdd"
		
		MesName = Get_SafeStr(Trim(Request("MesName")))
		MesTitle = Get_SafeStr(Trim(Request("MesTitle")))
		MesSex = Get_SafeStr(Trim(Request("MesSex")))
		MesContent = Get_SafeStr(Trim(Request("MesContent")))
		VerifyCode = Get_SafeStr(Trim(Request("VerifyCode")))
		
		MesContent=server.htmlencode(MesContent)
		MesContent=replace(MesContent,vbCrLf,"<br>")
		MesContent=replace(MesContent," ","&nbsp;")
		
		
		if session("VerifyCode")<>VerifyCode then
		response.write "<script language='JavaScript'>alert('��֤�����');" & "history.back()" & "</script>"
		Response.End()
		else
		session("VerifyCode")=""
		end if
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Message"
		rs.open sql,conn,1,3
		rs.addnew
		rs("MesName")=MesName
		rs("MesTitle")=MesTitle
		rs("MesSex")=MesSex
		rs("MesContent")=MesContent
		rs("MesIP")=Request.servervariables("remote_addr")
		rs("MesTime")=now()
		rs("SaveType")=0
		rs.update
		rs.close
		set rs=nothing
		%>
		<script language="javascript">
		alert('���������ѳɹ��ύ������Ա��˺󣬽���������վ�У�лл��');
		history.go(-1);
		</script>
		<%
		Response.End()


Case "CustomerReAdd"

		MesName = Get_SafeStr(Trim(Request("MesName")))
		MesTitle = Get_SafeStr(Trim(Request("MesTitle")))
		MesSex = Get_SafeStr(Trim(Request("MesSex")))
		MesContent = Get_SafeStr(Trim(Request("MesContent")))
		VerifyCode = Get_SafeStr(Trim(Request("VerifyCode")))
		
		MesContent=server.htmlencode(MesContent)
		MesContent=replace(MesContent,vbCrLf,"<br>")
		MesContent=replace(MesContent," ","&nbsp;")
		
		
		if session("VerifyCode")<>VerifyCode then
		response.write "<script language='JavaScript'>alert('��֤�����');" & "history.back()" & "</script>"
		Response.End()
		else
		session("VerifyCode")=""
		end if
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Message"
		rs.open sql,conn,1,3
		rs.addnew
		rs("MesName")=MesName
		rs("MesTitle")=MesTitle
		rs("MesSex")=MesSex
		rs("MesContent")=MesContent
		rs("MesIP")=Request.servervariables("remote_addr")
		rs("MesTime")=now()
		rs("SaveType")=1
		rs.update
		rs.close
		set rs=nothing
		%>
		<script language="javascript">
		alert('���������ѳɹ��ύ������Ա��˺󣬽���������վ�У�лл��');
		history.go(-1);
		</script>
		<%
		Response.End()


End Select
%>