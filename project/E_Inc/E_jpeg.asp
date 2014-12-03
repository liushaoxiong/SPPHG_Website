<% '缩略图
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
' 保存缩略图到指定文件夹下
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
' 注销实例
Set Jpeg = Nothing

JpegZoom = DirFileName
'JpegZoom = Ow&"|"&Nw
End Function



' 加水印
Sub JpegSign(SignImgUrl,SignColor,SignFont,SignContent)
Set Jpeg = Server.CreateObject("Persits.Jpeg")
' 打开目标图片
Jpeg.Open Server.MapPath(SignImgUrl)

' 添加文字水印
Jpeg.Canvas.Font.Color = &Hffffff
Jpeg.Canvas.Font.Size   =   24
Jpeg.Canvas.Font.Family = SignFont
Jpeg.Canvas.Font.Bold = False  
Jpeg.Canvas.Font.ShadowXoffset   =   1
Jpeg.Canvas.Font.ShadowYoffset   =   1
Jpeg.Canvas.Font.Quality = 4
Jpeg.Canvas.Print 10, 10,SignContent

' 保存文件
Jpeg.Save Server.MapPath(SignImgUrl)

' 注销对象
Set Jpeg = Nothing
end sub
%>