<%
'2007-2012 �������޹�˾ ��Ȩ���� ��������Ȩ��
'Time 2012-08-27 Program By Lqm

'��Ȩ�ͻ�ʹ�õ���ʼ���ڣ����ڰ�����������ǰһ�����
AuthorizedDate = "2012-10-15"

'��վ��Ȩ�ͻ�����
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