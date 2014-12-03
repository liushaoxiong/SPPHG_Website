<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Md5.asp"-->
<%
UserName=CheckContent(request.form("UserName"))
Passwords=Md5(CheckContent(request.form("Passwords")))
set rs = server.createobject("adodb.recordset")
sql="select * from E_Admin where MagName='"&UserName&"'"
rs.open sql,conn,1,3

if rs.eof then
   response.write "<script language=javascript> alert('管理员名称不正确，请重新输入。');location.replace('Login.asp');</script>"
   response.end
else
   AMagName=rs("MagName")
   APassword=rs("MagPass")
   MagLevel=rs("MagLevel")
   MagLock=rs("MagLock")
   MagPopdem=rs("MagPopdem")
end if

if APassword<>Passwords then
   response.write "<script language=javascript> alert('管理员密码不正确，请重新输入。');location.replace('Login.asp');</script>"
   response.end
end if 


if MagLock = 1 then
   response.write "<script language=javascript> alert('不能登录，此管理员帐号已被锁定。');location.replace('Login.asp');</script>"
   response.end
end if 
 
   rs("lstTime")=now()
   rs("lstIp")=Request.ServerVariables("Remote_Addr")
   rs.update
   rs.close
   set rs=nothing 

   session("AUserName")=AMagName
   session("MagLevel")=MagLevel
   session("MagPopdem")=MagPopdem
   session("LoginSystem")="Succeed"
   session.timeout=60

   response.redirect "E_ManageMain.asp"
   response.end

%>
