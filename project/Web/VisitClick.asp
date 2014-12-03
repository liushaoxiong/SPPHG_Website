<!--#include file="../E_Inc/Conn.asp"-->
<%

ID = Trim(Request.QueryString("ID"))

Set rs = Server.CreateObject("Adodb.Recordset")

Sql = "select ClickNum from E_News where ID=" & ID

rs.open sql,conn,1,3

rs("ClickNum") = rs("ClickNum") + 1

rs.update

hits = rs("ClickNum")

%>

document.write("<%=hits%>")
