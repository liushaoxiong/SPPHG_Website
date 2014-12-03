<%
'==================================================
'函数:		IfInt(IntX)
'参数:		需要验证的值
'作用:		判断参数是否为数字,否则函数等0
'==================================================
Dim IntX
Function FunIfInt(IntX)
	If not(IsNumeric(IntX)) or isnull(IntX) or IntX = "" then
		FunIfInt = 0
	Else
		FunIfInt = IntX
	End if
End Function
'==================================================
'函数:		StrLeft(Str,StrLen)
'参数:		字符串；需要截取的数量（汉字乘以2）
'作用:		将文字标题str（标题）根据lennum（字数）调出来
'==================================================
function StrLeft(Str,StrLen)
		dim L,T,I,C
		if Str="" then
		StrLeft=""
		exit function
		end if
		Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
		L=Len(Str)
		T=0
		for i=1 to L
		C=Abs(AscW(Mid(Str,i,1)))
		if C>255 then
		T=T+2
		else
		T=T+1
		end if
		if T>=StrLen then
		StrLeft=Left(Str,i) & ".."
		exit for
		else
		StrLeft=Str
		end if
		next
		StrLeft=Replace(Replace(Replace(replace(StrLeft," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'==================================================
'函数:		SafeStr(str)
'参数:		字符串
'作用:		得到安全的字符串
'==================================================
Dim Safestr
Function Get_SafeStr(Safestr)
	Get_SafeStr = Replace(Replace(Replace(Trim(Safestr), "'", ""), Chr(34), ""), ";", "")
End Function


'==================================================
' 过程:		SubDelFile(sPathFile)
' 参数: 	sPathFile:需要删除的文件路径
' 作用:		判断TemChangeStr是否为空，如果不为空，则在循环的记录里通过鼠标事件显示内容
'===================================================
Sub SubDelFile(sPathFile)
	On Error Resume Next
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	oFSO.DeleteFile(Server.MapPath(sPathFile))
	Set oFSO = Nothing
End Sub	

'==================================================
'函数:		CheckContent(byVal Str)
'参数:		需要过滤检查的字符串
'作用:		过滤可能威胁数据库的字符
'==================================================
Function CheckContent(ByVal Str)
		If IsNull(Str) Then
		UnCheckStr = ""
		Exit Function 
		End If
	    Str = Replace(Str,"&#39;","'")
        Str = Replace(Str,"&#34;","""")
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(w)(h&#101;re)"
	    str = re.replace(str,"$1here")
		re.Pattern="(s)(el&#101;ct)"
	    str = re.replace(str,"$1elect")
		re.Pattern="(i)(ns&#101;rt)"
	    str = re.replace(str,"$1nsert")
		re.Pattern="(c)(r&#101;ate)"
	    str = re.replace(str,"$1reate")
		re.Pattern="(d)(ro&#112;)"
	    str = re.replace(str,"$1rop")
		re.Pattern="(a)(lt&#101;r)"
	    str = re.replace(str,"$1lter")
		re.Pattern="(d)(el&#101;te)"
	    str = re.replace(str,"$1elete")
		re.Pattern="(u)(p&#100;ate)"
	    str = re.replace(str,"$1pdate")
		re.Pattern="(\s)(o&#114;)"
	    Str = re.replace(Str,"$1or")
		Set re=Nothing
        Str = Replace(Str, "&amp;", "&")
    	CheckContent=Str
End Function


' ===============================================
'将文本中的HTML标签转变为格式
' ===============================================
function changechr2(str2) 
changechr2=replace(replace(replace(replace(str2,"&lt;","<"),"&gt;",">"),"<br>",chr(13)),"&nbsp;"," ") 
end function

'==================================================
'过程:		FunListBox(SqlStr,FormName,OptText,OptValue,SelectText,DefaultSelected,FormCss)
'参数:		Sql语句；控件name；选项文字；选项值；已选值；默认值（一般为提示性说明）;控件样式
'作用:		得到下拉框,使用本函数需要在调用页面中使用myCloseObject(Conn)关闭数据库连接
'==================================================
Sub FunListBox(SqlStr,FormName,OptText,OptValue,SelectText,DefaultSelected,FormCss)
	set RsListBox = Conn.execute(SqlStr)
	Response.Write("<select name='"&FormName&"' id='"&FormName&"' class='"&FormCss&"'>")
	Response.write("<option value='0'>"&DefaultSelected&"</option>")	
	Do While Not RsListBox.eof
		Response.Write("<option value='"&RsListBox(OptValue)&"'")
		if SelectText <>"" and not isnull(SelectText) then
			If SelectText=RsListBox(OptValue) Then 
				Response.write (" Selected")
			End If
		End if
		Response.Write (">"&RsListBox(OptText)&"</option>")
		RsListBox.movenext
	loop
	Response.write("</select>")
	RsListBox.close
	set RsListBox =nothing
End Sub

'==================================================
'函数:		PictureUrl(ImgTrue)
'作用:		根据ImgUrlStr参数判断是否有图片
'==================================================
Function PictureUrl(ImgTrue)
	If isnull(ImgTrue) or ImgTrue="" then
		PictureUrl = "../E_img/nopic.gif"
	else
		PictureUrl = ImgTrue
	end if
End Function

'================================================
'函数名：SiteInfo
'作　用：调用网站基本信息
'参　数：
'返回值：定义的所有变量
'================================================
public SiteTitle,SiteUrl,SiteLogo,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,MesViewFlag,qq
sub SiteInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from Southidc_Site"
  rs.open sql,conn,1,1
  SiteTitle=rs("SiteTitle")
  Keywords=rs("Keywords")
  Descriptions=rs("Descriptions")
  SiteUrl=rs("SiteUrl")
  SiteLogo=rs("SiteLogo")
  ComName=rs("ComName")
  Address=rs("Address")
  ZipCode=rs("ZipCode")
  Telephone=rs("Telephone")
  Fax=rs("Fax")
  Email=rs("Email")
  qq=rs("qq")
  IcpNumber=rs("IcpNumber")
  MesViewFlag=rs("MesViewFlag")
  rs.close
  set rs=nothing
end sub

'==================================================
'函数:		GetExtendName(FileName)
'作用:		文件后缀
'==================================================
function GetExtendName(FileName)    
dim ExtName    
ExtName = LCase(FileName)    
ExtName = right(ExtName,3)    
ExtName = right(ExtName,3-Instr(ExtName,"."))    
GetExtendName = ExtName    
end function

'=========================================================
'调用广告
'=========================================================
Dim adid,width,height
sub Advertisement(adid,xwidth,xheight)
set rsad=server.CreateObject("adodb.recordset")
rsad.Open "Select * From E_Ads where id="&adid, conn,1,1
if rsad.eof then
response.write "<center>无广告</center>"
else

	AdImgType = GetExtendName(rsad("AdsContent"))
	select Case  AdImgType
	Case "swf"
	response.write "<object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 width="&xwidth&" height="&xheight&">"
	response.write "<param name=movie value="&rsad("AdsContent")&">"
	response.write "<param name=quality value=high>"
	response.write "<embed src="&rsad("AdsContent")&" quality=high pluginspage=http://www.macromedia.com/go/getflashplayer type=application/x-shockwave-flash width="&xwidth&" height="&xheight&"></embed>"
	response.write "</object>"
	Case Else
		response.write "<a href="&rsad("AdsLink")&" target=blank><img src="&rsad("AdsContent")&" height="&xheight&" width="&xwidth&"></a>"
	END select
end if
  rsad.close
  set rsad=nothing
end sub


'=========================================================
'调用广告
'=========================================================
sub Advertisement1(adid)
set rsad=server.CreateObject("adodb.recordset")
rsad.Open "Select * From E_Ads where id="&adid, conn,1,1
if rsad.eof then
response.write "<center>无广告</center>"
else
 If rsad("AdsContent") <> "" then
	AdImgType = GetExtendName(rsad("AdsContent"))
	End if
	If rsad("AdsContent") = "" then
	response.write "<center>&nbsp;</center>"
	Else
	select Case  AdImgType
	Case "swf"
	response.write "<object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0>"
	response.write "<param name=movie value="&rsad("AdsContent")&">"
	response.write "<param name=quality value=high>"
	response.write "<embed src="&rsad("AdsContent")&" quality=high pluginspage=http://www.macromedia.com/go/getflashplayer type=application/x-shockwave-flash width="&xwidth&" height="&xheight&"></embed>"
	response.write "</object>"
	Case Else
		response.write "<a href="&rsad("AdsLink")&" target=blank><img src="&rsad("AdsContent")&"></a>"
	END select
end if
End if
  rsad.close
  set rsad=nothing
end sub

'=========================================================
'调用文字友情链接的子程序,xCss指列表框的CSS样式类
'BAE9,B5E2,CDF7,C2E6,B0E5,C8A7,CBF8,D3CF,
'=========================================================
sub WordLink(Css)
response.write "<select class=input onChange=window.open(this.options[this.selectedIndex].value) name=select target=blank width=210>"
response.write "<option selected>友情链接</option>"
Set rsWordLink = Server.CreateObject("ADODB.Recordset")
sql ="Select * From E_FriendLink  Order By id ASC"
rsWordLink.open sql,Conn,1,1
do while not rsWordLink.eof
	response.write "<option value="&rsWordLink("LinkUrl")&">"&rsWordLink("LinkName")&"</option>"
	rsWordLink.MoveNext
Loop
rsWordLink.close
set rsWordLink=nothing
response.write "</select>"
end sub

'==================================================
'函数:		FunClassName(CId)
'作用:		根据栏目分类ID显示栏目名
'==================================================
Dim SId
Function FunClassName(SId)
SId = Int(SId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select Id,ClassName From E_NewsClass where id="&SId, Conn,1,1
	If RsClass.Eof Then
		FunClassName = ""
	Else
		FunClassName=RsClass("ClassName")
	End If
  RsClass.close
  set RsClass=nothing
End Function

'==================================================
'函数:		FunBClassName(TId)
'作用:		根据分类ID显示一级分类名称
'==================================================
Function FunBClassName(TId)
TId = Int(TId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select ID,BClassName From BClass where ID="&TId, Conn,1,1
	If RsClass.Eof Then
		FunBClassName = ""
	Else
		FunBClassName=RsClass("BClassName")
	End If
  RsClass.close
  set RsClass=nothing
End Function

'==================================================
'函数:		FunSClassName(CId)
'作用:		根据分类ID显示二级分类名称
'==================================================
Function FunSClassName(SId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select ID,SClassName From SClass where ID="&SId, Conn,1,1
	If RsClass.Eof Then
		FunSClassName = ""
	Else
		FunSClassName=RsClass("SClassName")
	End If
  RsClass.close
  set RsClass=nothing
End Function

'==================================================
'函数:		FunBrandName(CId)
'作用:		根据分类ID显示品牌名称
'==================================================
Function FunBrandName(SId)
SId = Int(SId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select ID,BrandName From E_Brand where ID="&SId, Conn,1,1
	If RsClass.Eof Then
		FunBrandName = ""
	Else
		FunBrandName=RsClass("BrandName")
	End If
  RsClass.close
  set RsClass=nothing
End Function


'==================================================
'函数:		FunPicFiles(TId)
'作用:		根据分类ID显示图片
'==================================================
Function FunPicFiles(TId)
TId = Int(TId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select ID,d_savefilename From E_product where ID="&TId, Conn,1,1
	If RsClass.Eof Then
		FunPicFiles = ""
	Else
		FunPicFiles=RsClass("d_savefilename")
	End If
  RsClass.close
  set RsClass=nothing
End Function

'==================================================
'函数:		FunAdsFiles(TId)
'作用:		根据分类ID显示图片
'==================================================
Function FunAdsFiles(TId)
TId = Int(TId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select ID,AdsContent From E_Ads where ID="&TId, Conn,1,1
	If RsClass.Eof Then
		FunAdsFiles = ""
	Else
		FunAdsFiles=RsClass("AdsContent")
	End If
  RsClass.close
  set RsClass=nothing
End Function

'==================================================
'函数:		FunAdsFiles(TId)
'作用:		根据分类ID显示图片
'==================================================
Function FunFrendlinkName(TId)
Select Case TID
Case 1
FunFrendlinkName="--国家省部委网站--"
Case 2
FunFrendlinkName="--全国各地党建网站--"
Case 3
FunFrendlinkName="--本地网站--"
End Select
End Function

'==================================================
'函数:		FunBClassName(TId)
'作用:		根据分类ID显示图片
'==================================================
Function FunPicFiles1(TId)
TId = Int(TId)
set RsClass=server.CreateObject("adodb.recordset")
RsClass.Open "Select ID,d_savefilename From E_News where ID="&TId, Conn,1,1
	If RsClass.Eof Then
		FunPicFiles1 = ""
	Else
		FunPicFiles1=RsClass("d_savefilename")
	End If
  RsClass.close
  set RsClass=nothing
End Function
%>
