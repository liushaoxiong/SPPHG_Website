<% '����ͼ
Response.cachecontrol="no-cache"
Response.addheader "pragma","no-cache"
Response.Expires=-1 

Function JpegZoom(FromImgUrl,ToWidth)

Set Jpeg = Server.CreateObject("Persits.Jpeg")
Path = Server.MapPath(FromImgUrl)

Jpeg.Open Path

if Jpeg.OriginalWidth > ToWidth then

	Jpeg.Width = ToWidth
	Jpeg.Height = ToWidth

	If Jpeg.OriginalWidth / Jpeg.OriginalHeight > 1 then
	Jpeg.Width = ToWidth 
	Jpeg.Height = int((ToWidth/Jpeg.OriginalWidth)*Jpeg.OriginalHeight)
	
	elseif Jpeg.OriginalWidth / Jpeg.OriginalHeight < 1 then
	Jpeg.Width = ToWidth
	Jpeg.Height= int((ToWidth/Jpeg.OriginalWidth)*Jpeg.OriginalHeight)
	
	end if

end if
' ��������ͼ��ָ���ļ�����
strExtension = Mid(FromImgUrl, InStrRev(FromImgUrl, ".")) 
randomize
ranNum=int(900*rnd)+100
filename=year(now())&month(now())&day(now())
			filename = filename &hour(now())&minute(now())
			filename = filename&second(now())&ranNum
			filename = filename&LCase(strExtension)
			
			sUploadDir = "/Upload/PicFiles/"
			sUploadDir =sUploadDir & Year(Now) & "-" & Month(Now)
		
			  Set objFSO = server.CreateObject("Scripting.FileSystemObject")
					If  objFSO.FolderExists(Server.MapPath(sUploadDir)) = False Then
					 objFSO.CreateFolder(Server.MapPath(sUploadDir))
					End If
			sUploadDir =sUploadDir &"/"
DirFileName = sUploadDir&filename
Jpeg.Save Server.MapPath(DirFileName)


'Ow = Jpeg.OriginalWidth
'Nw = ToWidth
' ע��ʵ��
Set Jpeg = Nothing

JpegZoom = DirFileName
'JpegZoom = Ow&"|"&Nw
End Function



' ��ˮӡ
Sub JpegSign(SignImgUrl,SignColor,SignFont,SignContent)
Set Jpeg = Server.CreateObject("Persits.Jpeg")
' ��Ŀ��ͼƬ
Jpeg.Open Server.MapPath(SignImgUrl)

' �������ˮӡ
Jpeg.Canvas.Font.Color = &Hffffff
Jpeg.Canvas.Font.Size   =   24
Jpeg.Canvas.Font.Family = SignFont
Jpeg.Canvas.Font.Bold = False  
Jpeg.Canvas.Font.ShadowXoffset   =   1
Jpeg.Canvas.Font.ShadowYoffset   =   1
Jpeg.Canvas.Font.Quality = 4
Jpeg.Canvas.Print 10, 10,SignContent

' �����ļ�
Jpeg.Save Server.MapPath(SignImgUrl)

' ע������
Set Jpeg = Nothing
end sub
%>