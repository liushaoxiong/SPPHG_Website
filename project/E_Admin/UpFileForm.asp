<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>文件选择</title>
<link href="Img/Main.css" rel="stylesheet" type="text/css">
</head> 
<body>
<table width="400" border="0" align="center" cellpadding="12" cellspacing="1" bgcolor="#CCCCCC">
  <form action="UpFileSave.asp?Result=<%=request.QueryString("Result")%>" method="post" enctype="multipart/form-data" name="formUpload">
  <tr>
    <td bgcolor="#FFFFFF">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="60" height="30" nowrap>选择文件：</td>
        <td><input name="FromFile" type="file" id="FromFile" size="41" class="input"></td>
      </tr>

      <tr>
        <td height="36" colspan="2" align="center" valign="bottom"><input name="reset" type="reset" class="button" value=" 重置 ">
          &nbsp;<input name="Submit" type="submit" class="button" value=" 上传 "></td>
        </tr>
    </table>
	</td>
  </tr>
  </form>
</table>
</body>
</html>

