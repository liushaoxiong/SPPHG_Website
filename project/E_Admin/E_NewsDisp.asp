<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->

<%
Default=2
ID = FunIfInt(Request("ID"))
Set Rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from E_News where  id="&ID
Rs.Open sql,conn,1,1
if Rs.eof and Rs.bof then
response.Write("û�м�¼")
else
ClassID = Rs("ClassID")
SortID = Rs("SortID")

'========�ж��Ƿ���й���Ȩ��
If session("MagLevel") <> 0 Then
if Instr(session("MagPopdem"),"|ClassID"&ClassID&",")=0 then 
  response.write "<script language='JavaScript'>alert('����Ȩ������ҳ�棬�뷵�أ�');" & "history.back()" & "</script>"
  response.end
end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������̨</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
<script language="javascript">
function ifmsgbox(message,url) {
	     if ( confirm(message) ) {
		 window.location.href=url
		 }
	}
</script>
</head>
<body>
<!--#include file="Top.html"-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" valign="top" bgcolor="f2f2f2">
    <!--#include file="E_SysMenu.asp"-->
    </td>
    <td width="205" valign="top" background="Img/C_bg.jpg">
    <!--#include file="E_NewsMenu.asp"-->
    </td>
    <td width="20"></td>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
      <table width="95%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FAFAFA">
        <tr>
          <td height="400" align="center" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="Img/menu_12.jpg">
            <tr>
              <td width="47"><img src="Img/menu_10.jpg" width="47" height="55"></td>
              <td class="Wryh T14 P2"><strong>&nbsp;�������</strong></td>
              <td width="73"><img src="Img/menu_14.jpg" width="73" height="55"></td>
            </tr>
          </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" background="Img/Add_banner.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="20">&nbsp;</td>
                    <td width="103" height="27" background="Img/qb.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="30">&nbsp;</td>
                          <td align="center" class="white Wryh T14"><a href="E_NewsMain.asp?ClassID=<%=ClassID%>"><strong>�����б�</strong></a></td>
                        </tr>
                    </table></td>
                    <td width="10">&nbsp;</td>
                    <td width="103" background="Img/tj1.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="30">&nbsp;</td>
                          <td align="center" class="white Wryh T14"><a href="E_NewsAdd.asp?ClassID=<%=ClassID%>"><strong>�������</strong></a></td>
                        </tr>
                    </table></td>
                     <td width="10">&nbsp;</td>
                    <td width="103" background="Img/Back.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="30">&nbsp;</td>
                        <td align="center" class="white Wryh T14"><a href="Javascript:history.back();"><strong>�����ϼ�</strong></a></td>
                      </tr>
                    </table></td> 
                    <td>&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
            </table>
           <div id="Empty"></div> 
           <table width="95%" border="0" cellspacing="0" cellpadding="0">
             <tr>
               <td height="40" class="T16 Wryh Dot"><%=Rs("Title")%></td>
             </tr>
             <tr>
               <td height="40" class="Dot">������Ŀ��<%=FunClassName(Rs("ClassID"))%>��|������ʱ�䣺<%=Rs("Addtime")%>��|���鿴������<%=Rs("ClickNum")%>��|��<a href="E_NewsEdit.asp?ID=<%=Rs("ID")%>">�޸�</a>��|��<a href="#" class="Dot" onClick="ifmsgbox('���������ݳ���ɾ����\n\nɾ�����ɻָ�������ص��ļ�Ҳ��һ��ɾ����','E_ManageSave.asp?SaveType=NewsDel&Id=<%=Rs("Id")%>&ClassId=<%=Rs("ClassId")%>&CurPage=1');">ɾ��</a></td>
             </tr>
             <tr>
               <td>&nbsp;</td>
             </tr>
             <tr>
               <td><%=Rs("Content")%></td>
             </tr>
           </table>
          </td>
        </tr>
      </table>
      <table width="95%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14"><img src="Img/Centet_10.jpg" width="14" height="13"></td>
          <td background="Img/Centet_11.jpg"><img src="Img/Centet_11.jpg"></td>
          <td width="7"><img src="Img/Centet_13.jpg" width="7" height="13"></td>
        </tr>
      </table>
      <table width="95%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="40">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Foot.html"-->
</body>
</html>
<%
end if
Rs.close
set Rs=nothing
%>