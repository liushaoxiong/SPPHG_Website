<%
'==================================================
'����:		IfInt(IntX)
'����:		��Ҫ��֤��ֵ
'����:		�жϲ����Ƿ�Ϊ����,��������0
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
'����:		StrLeft(Str,StrLen)
'����:		�ַ�������Ҫ��ȡ�����������ֳ���2��
'����:		�����ֱ���str�����⣩����lennum��������������
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
'����:		SafeStr(str)
'����:		�ַ���
'����:		�õ���ȫ���ַ���
'==================================================
Dim Safestr
Function Get_SafeStr(Safestr)
	Get_SafeStr = Replace(Replace(Replace(Trim(Safestr), "'", ""), Chr(34), ""), ";", "")
End Function


'==================================================
' ����:		SubDelFile(sPathFile)
' ����: 	sPathFile:��Ҫɾ�����ļ�·��
' ����:		�ж�TemChangeStr�Ƿ�Ϊ�գ������Ϊ�գ�����ѭ���ļ�¼��ͨ������¼���ʾ����
'===================================================
Sub SubDelFile(sPathFile)
	On Error Resume Next
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	oFSO.DeleteFile(Server.MapPath(sPathFile))
	Set oFSO = Nothing
End Sub	

'==================================================
'����:		CheckContent(byVal Str)
'����:		��Ҫ���˼����ַ���
'����:		���˿�����в���ݿ���ַ�
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
'���ı��е�HTML��ǩת��Ϊ��ʽ
' ===============================================
function changechr2(str2) 
changechr2=replace(replace(replace(replace(str2,"&lt;","<"),"&gt;",">"),"<br>",chr(13)),"&nbsp;"," ") 
end function

'==================================================
'����:		FunListBox(SqlStr,FormName,OptText,OptValue,SelectText,DefaultSelected,FormCss)
'����:		Sql��䣻�ؼ�name��ѡ�����֣�ѡ��ֵ����ѡֵ��Ĭ��ֵ��һ��Ϊ��ʾ��˵����;�ؼ���ʽ
'����:		�õ�������,ʹ�ñ�������Ҫ�ڵ���ҳ����ʹ��myCloseObject(Conn)�ر����ݿ�����
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
'����:		PictureUrl(ImgTrue)
'����:		����ImgUrlStr�����ж��Ƿ���ͼƬ
'==================================================
Function PictureUrl(ImgTrue)
	If isnull(ImgTrue) or ImgTrue="" then
		PictureUrl = "../E_img/nopic.gif"
	else
		PictureUrl = ImgTrue
	end if
End Function

'================================================
'��������SiteInfo
'�����ã�������վ������Ϣ
'�Ρ�����
'����ֵ����������б���
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
'����:		GetExtendName(FileName)
'����:		�ļ���׺
'==================================================
function GetExtendName(FileName)    
dim ExtName    
ExtName = LCase(FileName)    
ExtName = right(ExtName,3)    
ExtName = right(ExtName,3-Instr(ExtName,"."))    
GetExtendName = ExtName    
end function

'=========================================================
'���ù��
'=========================================================
Dim adid,width,height
sub Advertisement(adid,xwidth,xheight)
set rsad=server.CreateObject("adodb.recordset")
rsad.Open "Select * From E_Ads where id="&adid, conn,1,1
if rsad.eof then
response.write "<center>�޹��</center>"
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
'���ù��
'=========================================================
sub Advertisement1(adid)
set rsad=server.CreateObject("adodb.recordset")
rsad.Open "Select * From E_Ads where id="&adid, conn,1,1
if rsad.eof then
response.write "<center>�޹��</center>"
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
'���������������ӵ��ӳ���,xCssָ�б���CSS��ʽ��
'BAE9,B5E2,CDF7,C2E6,B0E5,C8A7,CBF8,D3CF,
'=========================================================
sub WordLink(Css)
response.write "<select class=input onChange=window.open(this.options[this.selectedIndex].value) name=select target=blank width=210>"
response.write "<option selected>��������</option>"
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
'����:		FunClassName(CId)
'����:		������Ŀ����ID��ʾ��Ŀ��
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
'����:		FunBClassName(TId)
'����:		���ݷ���ID��ʾһ����������
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
'����:		FunSClassName(CId)
'����:		���ݷ���ID��ʾ������������
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
'����:		FunBrandName(CId)
'����:		���ݷ���ID��ʾƷ������
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
'����:		FunPicFiles(TId)
'����:		���ݷ���ID��ʾͼƬ
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
'����:		FunAdsFiles(TId)
'����:		���ݷ���ID��ʾͼƬ
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
'����:		FunAdsFiles(TId)
'����:		���ݷ���ID��ʾͼƬ
'==================================================
Function FunFrendlinkName(TId)
Select Case TID
Case 1
FunFrendlinkName="--����ʡ��ί��վ--"
Case 2
FunFrendlinkName="--ȫ�����ص�����վ--"
Case 3
FunFrendlinkName="--������վ--"
End Select
End Function

'==================================================
'����:		FunBClassName(TId)
'����:		���ݷ���ID��ʾͼƬ
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
