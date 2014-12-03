<%
   if session("AUserName")=""  or session("MagLevel")="" or session("LoginSystem")<>"Succeed" then
     response.write "<script language=javascript> alert('ÄúÄ¿Ç°ÉÐÎ´µÇÂ½,ÇëµÇÂ½!');location.replace('Login.asp');</script>"
     response.end
  end if
%>