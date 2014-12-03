<div id="massage_box"><div class="massage">
<div class="header" onmousedown=MDown(massage_box)><div style="display:inline; width:150px; position:absolute; padding:5px 0 0 10px"><strong class="Wryh T14">添加大类</strong></div>
<span onClick="massage_box.style.visibility='hidden'; mask.style.visibility='hidden'" style="float:right; display:inline; cursor:hand; font-size:16px;padding:5px 10px 0 10px"><strong>×</strong></span></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
               <form name="editForm" method="post" action="E_ManageSave.asp?SaveType=BClassAdd" onSubmit="return checkdata()">
                 <tr>
                   <td width="100" height="40" align="center" class="Wryh T14">分类名称</td>
                   <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                         <td background="Img/Input_03.jpg"><input name="BClassName" class="Titleinput"></td>
                         <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                       </tr>
                   </table></td>
                 </tr>
                 <tr>
                   <td width="100" height="100" align="center" class="Wryh T14">分类简介</td>
                   <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td width="8"><img src="Img/input1_01.jpg" width="8" height="86"></td>
                         <td background="Img/Input1_03.jpg"><textarea name="Content" rows="4" class="Title1input"></textarea></td>
                         <td width="8"><img src="Img/input1_05.jpg" width="8" height="86"></td>
                       </tr>
                   </table></td>
                 </tr>
                 <tr>
                   <td width="100" height="40" align="center" class="Wryh T14">标题图片</td>
                   <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
                             <tr>
                               <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                               <td background="Img/Input_03.jpg"><input name="Pictures" class="Titleinput"></td>
                               <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                             </tr>
                         </table></td>
                         <td width="20%"><a href="javaScript:OpenScript('UpFileForm.asp?Result=Pictures',460,180)"><img src="Img/Upload.jpg" width="91" height="30" border="0" align="absmiddle"></a></td>
                         <td width="10%">&nbsp;</td>
                       </tr>
                   </table></td>
                 </tr>

                 <tr>
                   <td width="100" height="50" align="center">&nbsp;</td>
                   <td><input type="image" src="Img/Tijiao.jpg"></td>
                 </tr>
               </form>
          </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>
</div></div>
<div id="mask"></div>

<div id="massage_box1"><div class="massage1">
<div class="header" onmousedown=MDown(massage_box1)><div style="display:inline; width:150px; position:absolute; padding:5px 0 0 10px"><strong class="Wryh T14">添加小类</strong></div>
<span onClick="massage_box1.style.visibility='hidden'; mask1.style.visibility='hidden'" style="float:right; display:inline; cursor:hand; font-size:16px;padding:5px 10px 0 10px"><strong>×</strong></span></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
               <form name="editForm1" method="post" action="E_ManageSave.asp?SaveType=SClassAdd" onSubmit="return checkdata1()">
               <%If UpId = 0 Then%>
               <tr>
                   <td width="100" height="40" align="center" class="Wryh T14">所属大类</td>
                   <td><%call FunListBox("Select * from BClass where IfShow= 0 order by Id Asc","UpId","BClassName","Id",0,"请选择大类","input")%></td>
                 </tr>
                                  </tr>
                 <%Else%>
                 <input type="hidden" name="UpId" value="<%=UpId%>">
                 <%End If%>
                 <tr>
                   <td width="100" height="60" align="center" class="Wryh T14">分类名称</td>
                   <td><table width="90%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                         <td width="8"><img src="Img/Input_01.jpg" width="8" height="33"></td>
                         <td background="Img/Input_03.jpg"><input name="SClassName" class="Titleinput"></td>
                         <td width="8"><img src="Img/Input_05.jpg" width="8" height="33"></td>
                       </tr>
                   </table></td>

                 <tr>
                   <td width="100" height="50"></td>
                   <td align="center"><input type="image" src="Img/Tijiao.jpg"></td>
                 </tr>
               </form>
          </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>
</div></div>
<div id="mask1"></div>