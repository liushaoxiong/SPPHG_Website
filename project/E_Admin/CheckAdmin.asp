<%
   if session("AUserName")=""  or session("MagLevel")="" or session("LoginSystem")<>"Succeed" then
     response.write "<script language=javascript> alert('��Ŀǰ��δ��½,���½!');location.replace('Login.asp');</script>"
     response.end
  end if
%>