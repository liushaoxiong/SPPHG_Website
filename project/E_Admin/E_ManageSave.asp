<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Admin/CheckAdmin.asp"-->
<!--#include file="../E_Inc/E_Md5.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<!--#include file="E_htmlconfig.asp"-->


<%
SaveType = Get_SafeStr(Trim(Request("SaveType")))
MemID=session("UserID")

Select Case  SaveType

'添加企业信息
Case "AboutAdd"


		Title = CheckContent(Request.Form("Title"))
		Title1 = CheckContent(Request.Form("Title1"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		
		Title1=server.htmlencode(Title1)
		Title1=replace(Title1,vbCrLf,"<br>")
		Title1=replace(Title1," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_About"
		rs.open sql,conn,1,3
		rs.addnew
		rs("Title")=Title
		rs("Title1")=Title1
		rs("Pictures")=Pictures
		rs("Addtime")=Now()
		rs("Content")=Content
		rs.update
		rs.close
		set rs=nothing
		
		Newid=conn.execute("select max(Id) from E_About ")(0)
		Call htmll("","","/About/About-"&Newid&".Html","Web/Aboutus.asp","ID=",Newid,"","")
		response.Write "<script language=javascript>alert('添加成功！');location.href='E_AboutMain.asp';</script>"
		Response.End




'修改企业信息
Case "AboutEdit"


		ID = FunIfInt(Request("ID"))
		Title = CheckContent(Request.Form("Title"))
		Title1 = CheckContent(Request.Form("Title1"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		
		Title1=server.htmlencode(Title1)
		Title1=replace(Title1,vbCrLf,"<br>")
		Title1=replace(Title1," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_About where ID ="&ID
		rs.open sql,conn,1,3
		rs("Title")=Title
		rs("Title1")=Title1
		rs("Pictures")=Pictures
		rs("Addtime")=Now()
		rs("Content")=Content
		rs.update
		rs.close
		set rs=nothing
		
		Call htmll("","","/About/About-"&ID&".Html","Web/Aboutus.asp","ID=",ID,"","")
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_AboutMain.asp';</script>"
		Response.End

'配置网站信息
Case "Config"


		ClientTitle = CheckContent(Request.Form("ClientTitle"))
		ClientkeyWords = CheckContent(Request.Form("ClientkeyWords"))
		Clientdescription = CheckContent(Request.Form("Clientdescription"))
		ClientTelephone = CheckContent(Request.Form("ClientTelephone"))
		ClientMobile = CheckContent(Request.Form("ClientMobile"))
		ClientCompany = CheckContent(Request.Form("ClientCompany"))
		ClientOicq = CheckContent(Request.Form("ClientOicq"))
		ClientAddress = CheckContent(Request.Form("ClientAddress"))
		ClientFax = CheckContent(Request.Form("ClientFax"))
		Clientceo = CheckContent(Request.Form("Clientceo"))
		ClientICP = CheckContent(Request.Form("ClientICP"))
		ClientMail = CheckContent(Request.Form("ClientMail"))
		
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Config where ID =1"
		rs.open sql,conn,1,3
		rs("ClientTitle")=ClientTitle
		rs("ClientkeyWords")=ClientkeyWords
		rs("Clientdescription")=Clientdescription
		rs("ClientTelephone")=ClientTelephone
		rs("ClientMobile")=ClientMobile
		rs("ClientCompany")=ClientCompany
		rs("ClientOicq")=ClientOicq
		rs("ClientAddress")=ClientAddress
		rs("ClientFax")=ClientFax
		rs("Clientceo")=Clientceo
		rs("ClientICP")=ClientICP
		rs("ClientMail")=ClientMail
		rs.update
		rs.close
		set rs=nothing
		
		response.Write "<script language=javascript>alert('设置成功！');location.href='E_SetSite.asp';</script>"
		Response.End

'添加新闻
Case "NewsAdd"


		ClassID = FunIfInt(Request.Form("ClassID"))
		SortID = FunIfInt(Request.Form("SortID"))
		ScopeID = FunIfInt(Request.Form("ScopeID"))
		Title = CheckContent(Request.Form("Title"))
		Title1 = CheckContent(Request.Form("Title1"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		d_originalfilename = CheckContent(Request.Form("d_originalfilename"))
		d_savefilename = CheckContent(Request.Form("d_savefilename"))
		
		Title1=server.htmlencode(Title1)
		Title1=replace(Title1,vbCrLf,"<br>")
		Title1=replace(Title1," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_News"
		rs.open sql,conn,1,3
		rs.addnew
		rs("Title")=Title
		rs("Title1")=Title1
		rs("ClassID")=ClassID
		rs("ScopeID")=ScopeID
		rs("Pictures")=Pictures
		rs("SortID")=SortID
		rs("Addtime")=Now()
		rs("Content")=Content
		rs("d_originalfilename")=d_originalfilename
		rs("d_savefilename")=d_savefilename
		rs.update
		rs.close
		set rs=nothing
		
		Newid=conn.execute("select max(Id) from E_News ")(0)
		Call htmll("","","/News/E_NewsShow-"&Newid&"-"&ClassID&"-"&encode(Newid)&".Html","Web/Newsshow.asp","ID=",Newid,"","")
		response.Write "<script language=javascript>alert('添加成功！');location.href='E_NewsAdd.asp?ClassID="&ClassID&"';</script>"
		Response.End
		
		
'修改新闻
Case "NewsEdit"


		ID = FunIfInt(Request("ID"))
		ClassID = FunIfInt(Request("ClassID"))
		ScopeID = FunIfInt(Request.Form("ScopeID"))
		SortID = FunIfInt(Request("SortID"))
		Title = CheckContent(Request.Form("Title"))
		Title1 = CheckContent(Request.Form("Title1"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		d_originalfilename = CheckContent(Request.Form("d_originalfilename"))
		d_savefilename = CheckContent(Request.Form("d_savefilename"))
		
		Title1=server.htmlencode(Title1)
		Title1=replace(Title1,vbCrLf,"<br>")
		Title1=replace(Title1," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_News where ID ="&ID
		rs.open sql,conn,1,3
		rs("Title")=Title
		rs("Title1")=Title1
		rs("ScopeID")=ScopeID
		rs("Pictures")=Pictures
		rs("SortID")=SortID
		rs("Addtime")=Now()
		rs("Content")=Content
		rs("d_originalfilename")=d_originalfilename
		rs("d_savefilename")=d_savefilename
		rs.update
		rs.close
		set rs=nothing
		
		Call htmll("","","/News/E_NewsShow-"&ID&"-"&ClassID&"-"&encode(ID)&".Html","Web/Newsshow.asp","ID=",ID,"","")
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_NewsMain.asp?ClassID="&ClassID&"';</script>"
		Response.End
		
		
'删除新闻		
Case "NewsDel"

		ID = FunIfInt(Request("ID"))
		ClassId = FunIfInt(Request("ClassId"))
		CurPage = FunIfInt(Request("CurPage"))
		Conn.Execute("Delete from E_News Where Id = "&ID)
		
		sSavePathFileName = FunPicFiles1(ID)
		'删除本记录上传过的图片等文件
		If sSavePathFileName <> "" Then
  		aSavePathFileName = Split(sSavePathFileName, "|")
		
  		Dim j
  		For j = 0 To UBound(aSavePathFileName)
		' 按路径文件名删除文件
		Call SubDelFile("/E_eWebEditor/uploadfile/"&aSavePathFileName(j))
  		Next
		End If
		
		Response.Redirect("E_NewsMain.asp?ClassId="&ClassId&"&CurPage="&CurPage&"")
		Response.End
		
		
		
'添加产品
Case "ProAdd"


		SortID = FunIfInt(Request.Form("SortID"))
		BClass = FunIfInt(Request.Form("BClass"))
		SClass = FunIfInt(Request.Form("SClass"))
		Price1 = FunIfInt(Request.Form("Price1"))
		Price2 = FunIfInt(Request.Form("Price2"))
		NewFlag = FunIfInt(Request.Form("NewFlag"))
		Recommend = FunIfInt(Request.Form("Recommend"))
		IfBuy = FunIfInt(Request.Form("IfBuy"))
		Brand = FunIfInt(Request.Form("Brand"))
		ProName = CheckContent(Request.Form("ProName"))
		Model = CheckContent(Request.Form("Model"))
		Title = CheckContent(Request.Form("Title"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Pictures1 = CheckContent(Request.Form("Pictures1"))
		Content = CheckContent(Request.Form("Content"))
		d_originalfilename = CheckContent(Request.Form("d_originalfilename"))
		d_savefilename = CheckContent(Request.Form("d_savefilename"))

		
		Title=server.htmlencode(Title)
		Title=replace(Title,vbCrLf,"<br>")
		Title=replace(Title," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_product"
		rs.open sql,conn,1,3
		rs.addnew
		rs("SortID")=SortID
		rs("BClass")=BClass
		rs("SClass")=SClass
		rs("Price1")=Price1
		rs("Price2")=Price2
		rs("NewFlag")=NewFlag
		rs("Recommend")=Recommend
		rs("IfBuy")=IfBuy
		rs("Brand")=Brand
		rs("ProName")=ProName
		rs("Model")=Model
		rs("Title")=Title
		rs("Pictures")=Pictures
		rs("Pictures1")=Pictures1
		rs("Addtime")=Now()
		rs("Content")=Content
		rs("d_originalfilename")=d_originalfilename
		rs("d_savefilename")=d_savefilename
		rs.update
		rs.close
		set rs=nothing
		
		Newid=conn.execute("select max(Id) from E_product ")(0)
		Call htmll("","","/Pro/E_ProShow-"&Newid&"-"&BClass&".Html","Web/ProShow.asp","ID=",Newid,"","")
		response.Write "<script language=javascript>alert('添加成功！');location.href='E_ProductAdd.asp';</script>"
		Response.End
		
		
'修改产品
Case "ProEdit"

		ID = FunIfInt(Request("ID"))
		SortID = FunIfInt(Request.Form("SortID"))
		BClass = FunIfInt(Request.Form("BClass"))
		SClass = FunIfInt(Request.Form("SClass"))
		Price1 = FunIfInt(Request.Form("Price1"))
		Price2 = FunIfInt(Request.Form("Price2"))
		NewFlag = FunIfInt(Request.Form("NewFlag"))
		Recommend = FunIfInt(Request.Form("Recommend"))
		IfBuy = FunIfInt(Request.Form("IfBuy"))
		Brand = FunIfInt(Request.Form("Brand"))
		ProName = CheckContent(Request.Form("ProName"))
		Model = CheckContent(Request.Form("Model"))
		Title = CheckContent(Request.Form("Title"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Pictures1 = CheckContent(Request.Form("Pictures1"))
		Content = CheckContent(Request.Form("Content"))
		d_originalfilename = CheckContent(Request.Form("d_originalfilename"))
		d_savefilename = CheckContent(Request.Form("d_savefilename"))

		
		Title=server.htmlencode(Title)
		Title=replace(Title,vbCrLf,"<br>")
		Title=replace(Title," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_product Where Id="&ID
		rs.open sql,conn,1,3
		rs("SortID")=SortID
		rs("BClass")=BClass
		rs("SClass")=SClass
		rs("Price1")=Price1
		rs("Price2")=Price2
		rs("NewFlag")=NewFlag
		rs("Recommend")=Recommend
		rs("IfBuy")=IfBuy
		rs("Brand")=Brand
		rs("ProName")=ProName
		rs("Model")=Model
		rs("Title")=Title
		rs("Pictures")=Pictures
		rs("Pictures1")=Pictures1
		rs("Addtime")=Now()
		rs("Content")=Content
		rs("d_originalfilename")=d_originalfilename
		rs("d_savefilename")=d_savefilename
		rs.update
		rs.close
		set rs=nothing
		Call htmll("","","/Pro/E_ProShow-"&ID&"-"&BClass&".Html","Web/ProShow.asp","ID=",ID,"","")
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_ProductMain.asp';</script>"
		Response.End
		

'删除产品		
Case "ProDel"


		ID = FunIfInt(Request("ID"))
		CurPage = FunIfInt(Request("CurPage"))
		
		sSavePathFileName = FunPicFiles(ID)
		'删除本记录上传过的图片等文件
		If sSavePathFileName <> "" Then
  		aSavePathFileName = Split(sSavePathFileName, "|")

  		For j = 0 To UBound(aSavePathFileName)
		' 按路径文件名删除文件
		Call SubDelFile("/E_eWebEditor/uploadfile/"&aSavePathFileName(j))
  		Next
		End If
		Conn.Execute("Delete from E_product Where Id ="&ID)
		Response.Redirect("E_ProductMain.asp?CurPage="&CurPage&"")
		Response.End
		

'全选删除产品		
Case "ProDelAll"


		ID = Trim(Request("ID"))
		CurPage = FunIfInt(Request("CurPage"))
		
		sSavePathFileName = FunPicFiles(ID)
		'删除本记录上传过的图片等文件
		If sSavePathFileName <> "" Then
  		aSavePathFileName = Split(sSavePathFileName, "|")

  		For j = 0 To UBound(aSavePathFileName)
		' 按路径文件名删除文件
		Call SubDelFile("/E_eWebEditor/uploadfile/"&aSavePathFileName(j))
  		Next
		End If
		Conn.Execute("Delete from E_product Where Id in ("&ID&")")
		Response.Redirect("E_ProductMain.asp?CurPage="&CurPage&"")
		Response.End
		
		
'添加产品品牌
Case "BrandAdd"

		
		BrandName = CheckContent(Request.Form("BrandName"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		
		Content=server.htmlencode(Content)
		Content=replace(Content,vbCrLf,"<br>")
		Content=replace(Content," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Brand"
		rs.open sql,conn,1,3
		rs.addnew
		rs("BrandName")=BrandName
		rs("Pictures")=Pictures
		rs("Content")=Content
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>location.href='E_BrandMain.asp';</script>"
		Response.End	
		
	
'修改产品品牌
Case "BrandEdit"

		ID = FunIfInt(Request("ID"))
		BrandName = CheckContent(Request.Form("BrandName"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		
		Content=server.htmlencode(Content)
		Content=replace(Content,vbCrLf,"<br>")
		Content=replace(Content," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Brand Where ID="&ID
		rs.open sql,conn,1,3
		rs("BrandName")=BrandName
		rs("Pictures")=Pictures
		rs("Content")=Content
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_BrandMain.asp';</script>"
		Response.End
		
'删除产品大类		
Case "BrandDel"

		ID = FunIfInt(Request("ID"))
		Conn.Execute("Delete from E_Brand Where Id = "&ID)
		Response.Redirect("E_BrandMain.asp")
		Response.End
	
			
		
'添加产品大类
Case "BClassAdd"

		
		BClassName = CheckContent(Request.Form("BClassName"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		
		Content=server.htmlencode(Content)
		Content=replace(Content,vbCrLf,"<br>")
		Content=replace(Content," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from BClass"
		rs.open sql,conn,1,3
		rs.addnew
		rs("BClassName")=BClassName
		rs("Pictures")=Pictures
		rs("Content")=Content
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>location.href='E_ClassMain.asp';</script>"
		Response.End
		
'修改产品大类
Case "BClassEdit"

		ID = FunIfInt(Request("ID"))
		BClassName = CheckContent(Request.Form("BClassName"))
		Pictures = CheckContent(Request.Form("Pictures"))
		Content = CheckContent(Request.Form("Content"))
		
		Content=server.htmlencode(Content)
		Content=replace(Content,vbCrLf,"<br>")
		Content=replace(Content," ","&nbsp;")
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from BClass Where ID="&ID
		rs.open sql,conn,1,3
		rs("BClassName")=BClassName
		rs("Pictures")=Pictures
		rs("Content")=Content
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_ClassMain.asp';</script>"
		Response.End
		
'修改产品大类状态
Case "IfShowEdit"


		ID = FunIfInt(Request("ID"))
		IfShow = FunIfInt(Request("IfShow"))
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from BClass where ID ="&ID
		rs.open sql,conn,1,3
		rs("IfShow")=IfShow
		rs.update
		rs.close
		set rs=nothing
		
		response.Write "<script language=javascript>location.href='E_ClassMain.asp';</script>"
		Response.End
		
'删除产品大类		
Case "BClassDel"

		ID = FunIfInt(Request("ID"))
		Conn.Execute("Delete from BClass Where Id = "&ID)
		Conn.Execute("Delete from SClass Where UpId = "&ID)
		Conn.Execute("Delete from E_product Where BClass = "&ID)
		Response.Redirect("E_ClassMain.asp")
		Response.End
		
'添加产品小类
Case "SClassAdd"

		UpId = FunIfInt(Request.Form("UpId"))
		SClassName = CheckContent(Request.Form("SClassName"))
		Pictures = CheckContent(Request.Form("Pictures"))
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from SClass"
		rs.open sql,conn,1,3
		rs.addnew
		rs("UpId")=UpId
		rs("SClassName")=SClassName
		rs("Pictures")=Pictures
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>location.href='E_SClassMain.asp?UpId="&UpId&"';</script>"
		Response.End
		
'修改产品小类
Case "SClassEdit"


		ID = FunIfInt(Request("ID"))
		UpId = FunIfInt(Request("UpId"))
	    SClassName = CheckContent(Request.Form("SClassName"))
		set rs = server.createobject("adodb.recordset")
		sql="select * from SClass where ID ="&ID
		rs.open sql,conn,1,3
		rs("SClassName")=SClassName
		rs.update
		rs.close
		set rs=nothing
		
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_SClassMain.asp?UpId="&UpId&"';</script>"
		Response.End
		
		
'删除产品小类		
Case "SClassDel"

		ID = FunIfInt(Request("ID"))
		UpId = FunIfInt(Request("UpId"))
		Conn.Execute("Delete from SClass Where ID = "&ID)
		Conn.Execute("Delete from E_product Where SClass = "&ID)
		response.Write "<script language=javascript>location.href='E_SClassMain.asp?UpId="&UpId&"';</script>"
		Response.End
		

'添加管理员
Case "AdminAdd"


		MagName = CheckContent(Request.Form("MagName"))
		
		if not conn.execute("select MagName from E_Admin where MagName='" & MagName & "'").eof then 
		response.write "<script language='JavaScript'>alert('此用户已存在，请返回修改！');" & "history.back()" & "</script>"
		response.end
		End if
		
		MagPass = CheckContent(Request.Form("MagPass"))
		MagLevel = FunIfInt(Request.Form("MagLevel"))
		Purview1 = CheckContent(Request.Form("Purview1"))
		Purview2 = CheckContent(Request.Form("Purview2"))
		Purview3 = CheckContent(Request.Form("Purview3"))
		Purview4 = CheckContent(Request.Form("Purview4"))
		Purview5 = CheckContent(Request.Form("Purview5"))
		Purview6 = CheckContent(Request.Form("Purview6"))
		Purview7 = CheckContent(Request.Form("Purview7"))
		Purview8 = CheckContent(Request.Form("Purview8"))
		Purview9 = CheckContent(Request.Form("Purview9"))
		Purview10 = CheckContent(Request.Form("Purview10"))
		Purview11 = CheckContent(Request.Form("Purview11"))
		Purview12 = CheckContent(Request.Form("Purview12"))
		Purview13 = CheckContent(Request.Form("Purview13"))
		Purview14 = CheckContent(Request.Form("Purview14"))
		Purview15 = CheckContent(Request.Form("Purview15"))
		Purview16 = CheckContent(Request.Form("Purview16"))
		ClassID = CheckContent(Request.Form("ClassID"))
'超级管理员		
If MagLevel = 0 Then

		Set rs = Server.CreateObject("ADODB.Recordset")
		SqlRs="select * from E_NewsClass where IfShow = 1"
		rs.Open SqlRs, Conn, 1										 
		For iRs = 1 to rs.Recordcount
		if rs.EOF then     
		Exit For 
		end if
		AClassID = AClassID&"|ClassID"&Rs("Id")&","
		rs.MoveNext 
		next
		rs.close
		set rs=nothing
		
		MagPopdem =	"|1,|2,|3,|4,|5,|6,|7,|8,|9,|10,|11,|12,|13,|14,|15,|16,"				
		MagPopdem = MagPopdem&AClassID	
Else

'普通管理员	
		If ClassID <> "" Then
		MagPopdem=Purview1&Purview2&Purview3&Purview4&Purview5&Purview6&Purview7&Purview8&Purview9&Purview10&Purview11&Purview12&Purview13&Purview14&Purview15&Purview16&ClassID&","
		Else
		MagPopdem=Purview1&Purview2&Purview3&Purview4&Purview5&Purview6&Purview7&Purview8&Purview9&Purview10&Purview11&Purview12&Purview13&Purview14&Purview15&Purview16
		End if
End If		

		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Admin"
		rs.open sql,conn,1,3
		rs.addnew
		rs("MagName")=MagName
		rs("MagPass")=Md5(MagPass)
		rs("MagLevel")=MagLevel
		rs("MagPopdem")=MagPopdem
		rs("lstTime")=Now()
		rs("lstIp")=Request.ServerVariables("REMOTE_ADDR") 
		rs.update
		rs.close
		set rs=nothing
		
		response.Write "<script language=javascript>alert('添加成功！');location.href='E_AdminUserMain.asp';</script>"
		Response.End
		
'修改管理员
Case "AdminEdit"


		ID = FunIfInt(Request("ID"))
		MagName = CheckContent(Request.Form("MagName"))
		MagPass = CheckContent(Request.Form("MagPass"))
		MagPass1 = CheckContent(Request.Form("MagPass1"))
		
		If MagPass <> MagPass1 then
		response.write "<script language='JavaScript'>alert('2次密码输入不同，请返回修改！');" & "history.back()" & "</script>"
		End If
		
		MagLevel = FunIfInt(Request.Form("MagLevel"))
		Purview1 = CheckContent(Request.Form("Purview1"))
		Purview2 = CheckContent(Request.Form("Purview2"))
		Purview3 = CheckContent(Request.Form("Purview3"))
		Purview4 = CheckContent(Request.Form("Purview4"))
		Purview5 = CheckContent(Request.Form("Purview5"))
		Purview6 = CheckContent(Request.Form("Purview6"))
		Purview7 = CheckContent(Request.Form("Purview7"))
		Purview8 = CheckContent(Request.Form("Purview8"))
		Purview9 = CheckContent(Request.Form("Purview9"))
		Purview10 = CheckContent(Request.Form("Purview10"))
		Purview11 = CheckContent(Request.Form("Purview11"))
		Purview12 = CheckContent(Request.Form("Purview12"))
		Purview13 = CheckContent(Request.Form("Purview13"))
		Purview14 = CheckContent(Request.Form("Purview14"))
		Purview15 = CheckContent(Request.Form("Purview15"))
		Purview16 = CheckContent(Request.Form("Purview16"))
		ClassID = CheckContent(Request.Form("ClassID"))
'超级管理员		
If MagLevel = 0 Then

		Set rs = Server.CreateObject("ADODB.Recordset")
		SqlRs="select * from E_NewsClass where IfShow = 1"
		rs.Open SqlRs, Conn, 1										 
		For iRs = 1 to rs.Recordcount
		if rs.EOF then     
		Exit For 
		end if
		AClassID = AClassID&"|ClassID"&Rs("Id")&","
		rs.MoveNext 
		next
		rs.close
		set rs=nothing
		
		MagPopdem =	"|1,|2,|3,|4,|5,|6,|7,|8,|9,|10,|11,|12,|13,|14,|15,|16,"				
		MagPopdem = MagPopdem&AClassID	
Else

'普通管理员	
		If ClassID <> "" Then
		MagPopdem=Purview1&Purview2&Purview3&Purview4&Purview5&Purview6&Purview7&Purview8&Purview9&Purview10&Purview11&Purview12&Purview13&Purview14&Purview15&Purview16&ClassID&","
		Else
		MagPopdem=Purview1&Purview2&Purview3&Purview4&Purview5&Purview6&Purview7&Purview8&Purview9&Purview10&Purview11&Purview12&Purview13&Purview14&Purview15&Purview16
		End if
End If
	    SClassName = CheckContent(Request.Form("SClassName"))
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Admin where ID ="&ID
		rs.open sql,conn,1,3
		rs("MagName")=MagName
		If MagPass <> "" Then
		rs("MagPass")=Md5(MagPass)
		End If
		rs("MagLevel")=MagLevel
		If MagLevel <> 0 Then
		rs("MagPopdem")=MagPopdem
		End If
		rs.update
		rs.close
		set rs=nothing
		
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_AdminUserMain.asp';</script>"
		Response.End
		
'删除管理员	
Case "AdminDel"

		ID = FunIfInt(Request("ID"))
		CurPage = FunIfInt(Request("CurPage"))
		Conn.Execute("Delete from E_Admin Where ID = "&ID)
		response.Write "<script language=javascript>location.href='E_AdminUserMain.asp?CurPage="&CurPage&"';</script>"
		Response.End
		
		
Case "DataBackup"'数据备份

		set Fso=Server.CreateObject("Scripting.FileSystemObject")
		Fso.CopyFile Server.MapPath(DataBaseFile),Server.MapPath(DataBaseBakFile)
		Set Fso=nothing
		response.Write "<script language=javascript>alert('数据备份成功！');location.href='E_DataBackup.asp';</script>"
		Response.End	


Case "DataResume"'数据恢复

		set Fso=Server.CreateObject("Scripting.FileSystemObject")
		Fso.CopyFile Server.MapPath(DataBaseBakFile),Server.MapPath(DataBaseFile)
		Set Fso=nothing
		response.Write "<script language=javascript>alert('数据恢复成功！');location.href='E_DataResume.asp';</script>"
		Response.End
		
'添加友情链接
Case "LinkAdd"

		LinkType = FunIfInt(Request.Form("LinkType"))
		ClassID = FunIfInt(Request.Form("ClassID"))
		LinkName = CheckContent(Request.Form("LinkName"))
		LinkUrl = CheckContent(Request.Form("LinkUrl"))
		Pictures = CheckContent(Request.Form("Pictures"))
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_FriendLink"
		rs.open sql,conn,1,3
		rs.addnew
		rs("LinkType")=LinkType
		rs("ClassID")=ClassID
		rs("LinkName")=LinkName
		rs("LinkUrl")=LinkUrl
		rs("Pictures")=Pictures
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('添加成功！');location.href='E_FriendlinkAdd.asp';</script>"
		Response.End
		
'修改友情链接
Case "LinkEdit"

		ID = FunIfInt(Request("ID"))
		LinkType = FunIfInt(Request.Form("LinkType"))
		ClassID = FunIfInt(Request.Form("ClassID"))
		LinkName = CheckContent(Request.Form("LinkName"))
		LinkUrl = CheckContent(Request.Form("LinkUrl"))
		Pictures = CheckContent(Request.Form("Pictures"))
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_FriendLink Where ID = "&ID
		rs.open sql,conn,1,3
		rs("LinkType")=LinkType
		rs("ClassID")=ClassID
		rs("LinkName")=LinkName
		rs("LinkUrl")=LinkUrl
		rs("Pictures")=Pictures
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_FriendlinkMain.asp';</script>"
		Response.End
		

'删除友情链接	
Case "LinkDel"

		ID = FunIfInt(Request("ID"))
		CurPage = FunIfInt(Request("CurPage"))
		Conn.Execute("Delete from E_FriendLink Where ID = "&ID)
		response.Write "<script language=javascript>location.href='E_FriendlinkMain.asp?CurPage="&CurPage&"';</script>"
		Response.End
		
'回复留言
Case "MesEdit"

		ID = FunIfInt(Request("ID"))
		IfShow = FunIfInt(Request.Form("IfShow"))
		ReContent = CheckContent(Request.Form("ReContent"))

		ReContent=server.htmlencode(ReContent)
		ReContent=replace(ReContent,vbCrLf,"<br>")
		ReContent=replace(ReContent," ","&nbsp;")

		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Message Where ID = "&ID
		rs.open sql,conn,1,3
		rs("IfShow")=IfShow
		rs("ReContent")=ReContent
		rs("ReTime")=Now()
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('留言回复成功！');history.go(-2);</script>"
		Response.End
		
'删除留言	
Case "MesDel"

		ID = FunIfInt(Request("ID"))
		CurPage = FunIfInt(Request("CurPage"))
		Conn.Execute("Delete from E_Message Where ID = "&ID)
        response.Write "<script language=javascript>history.go(-2);</script>"
		Response.End		

'添加广告
Case "AdsAdd"

		AdsTitle = CheckContent(Request.Form("AdsTitle"))
		AdsLink = CheckContent(Request.Form("AdsLink"))
		AdsContent = CheckContent(Request.Form("AdsContent"))
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Ads"
		rs.open sql,conn,1,3
		rs.addnew
		rs("AdsTitle")=AdsTitle
		rs("AdsLink")=AdsLink
		rs("AdsContent")=AdsContent
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('添加成功！');location.href='E_AdsMain.asp';</script>"
		Response.End
		
'修改广告
Case "AdsEdit"

		ID = FunIfInt(Request("ID"))
		AdsTitle = CheckContent(Request.Form("AdsTitle"))
		AdsLink = CheckContent(Request.Form("AdsLink"))
		AdsContent = CheckContent(Request.Form("AdsContent"))
		
		set rs = server.createobject("adodb.recordset")
		sql="select * from E_Ads Where ID = "&ID
		rs.open sql,conn,1,3
		rs("AdsTitle")=AdsTitle
		rs("AdsLink")=AdsLink
		rs("AdsContent")=AdsContent
		rs.update
		rs.close
		set rs=nothing
		response.Write "<script language=javascript>alert('修改成功！');location.href='E_AdsMain.asp';</script>"
		Response.End
		

'删除广告	
Case "AdsDel"

		ID = FunIfInt(Request("ID"))
		CurPage = FunIfInt(Request("CurPage"))
		
		' 按路径文件名删除文件
		aSavePathFileName=FunAdsFiles(ID)
		If aSavePathFileName <> "" Then
		Call SubDelFile(aSavePathFileName)
		End If
		
		Conn.Execute("Delete from E_Ads Where ID = "&ID)

		
		response.Write "<script language=javascript>location.href='E_AdsMain.asp?CurPage="&CurPage&"';</script>"
		Response.End


End Select
%>