<%
'2007-2012 网络有限公司 版权所有 保留所有权利
'Time 2012-08-27 Program By Lqm

'授权客户使用的起始日期，日期按交付日期提前一天计算
AuthorizedDate = "2012-10-15"

'网站授权客户资料
set rs = server.createobject("adodb.recordset")
sql="select * from E_Config where ID =1"
rs.open sql,conn,1,1
ClientTitle=rs("ClientTitle")
ClientkeyWords=rs("ClientkeyWords")
Clientdescription=rs("Clientdescription")
ClientTelephone=rs("ClientTelephone")
ClientMobile=rs("ClientMobile")
ClientCompany=rs("ClientCompany")
ClientOicq=rs("ClientOicq")
ClientAddress=rs("ClientAddress")
ClientFax=rs("ClientFax")
Clientceo=rs("Clientceo")
ClientICP=rs("ClientICP")
ClientMail=rs("ClientMail")
rs.close
set rs=nothing
%>