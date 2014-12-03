<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->
<!--#include file="../E_Inc/E_Md5.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->


<%

		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_NewsClass where ID =33"
		rs.open sql,conn,1,3
		rs("ClassName")="ÊÖÊõÊÓÆµ"

		rs.update
		rs.close
		set rs=nothing


%>