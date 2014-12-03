
<% 
Dim Fy_Url,Fy_a,Fy_x,Fy_Cs(),Fy_Cl,Fy_Ts,Fy_Zx 
'---定义部份 头------ 
Fy_Cl = 1 '处理方式：1=提示信息,2=转向页面,3=先提示再转向 
Fy_Zx = "index.Asp" '出错时转向的页面 
'---定义部份 尾------ 

On Error Resume Next 
Fy_Url=Request.ServerVariables("QUERY_STRING") 
Fy_a=split(Fy_Url,"&") 
redim Fy_Cs(ubound(Fy_a)) 
On Error Resume Next 
for Fy_x=0 to ubound(Fy_a) 
Fy_Cs(Fy_x) = left(Fy_a(Fy_x),instr(Fy_a(Fy_x),"=")-1) 
Next 
For Fy_x=0 to ubound(Fy_Cs) 
If Fy_Cs(Fy_x)<>"" Then 
If Instr(LCase(Request(Fy_Cs(Fy_x))),"'")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"and")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"select")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"update")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"chr")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"delete%20from")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),";")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"insert")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"mid")<>0 Or Instr(LCase(Request(Fy_Cs(Fy_x))),"master.")<>0 Then 
Select Case Fy_Cl 
Case "1" 
Response.Write "<Script Language=JavaScript>alert(' 出现错误！参数 "&Fy_Cs(Fy_x)&" 的值中包含非法字符串！\n\n 请不要在参数中出现：and,select,update,insert,delete,chr 等非法字符！\n\n我已经设置了不能SQL注入，请不要对我进行非法手段！');window.close();</Script>" 
Case "2" 
Response.Write "<Script Language=JavaScript>location.href='"&Fy_Zx&"'</Script>" 
Case "3" 
Response.Write "<Script Language=JavaScript>alert(' 出现错误！参数 "&Fy_Cs(Fy_x)&"的值中包含非法字符串！\n\n 请不要在参数中出现：,and,select,update,insert,delete,chr 等非法字符！\n\n设计了门，非法侵入请离开，谢谢！');location.href='"&Fy_Zx&"';</Script>" 
End Select 
Response.End 
End If 
End If 
Next 
%> 


<%
Dim DataBaseFile,DataBaseBakFile,DataPassword,UpFiles,Conn,ServerMapPath
DataBaseFileFolder = "../E_DadaBase"					    '数据库目录		（Access库）
DataBaseFile = "../E_DadaBase/1.mdb"		'数据库文件		（Access库）
DataBaseBakFile = "../E_DadaBase/EskyingDataBaseBak.asp"	'数据库备份文件	（Access库）
DataPassword = ""					        '数据库密码		（Access库）

'==================================================
'ACCESS数据库连接字符串
'==================================================
set Conn=Server.Createobject("Adodb.Connection")
ServerMapPath=Server.MapPath(DataBaseFile)
Conn.Open "Driver={Microsoft Access Driver (*.mdb)}; password="&DataPassword&"; DBQ=" & ServerMapPath

'==================================================
'SQLSERVER数据库连接字符串
'==================================================
'set Conn=server.createobject("adodb.connection")
'Conn.Open "driver={sql server};server=[local];database=Edata;uid=Sa;pwd=111111"
%>