<table width="100%" border="0" cellpadding="0" cellspacing="0" background="Img/Search.jpg">
 <Form name="Form" method="post" action="E_ProductMain.asp">   
      <tr>
      <td width="18" height="61">&nbsp;</td>
      <td height="61"><input name="Key" style="width:100%; background:none; border:none; color:#fff; padding-top:2px" onFocus="if(this.value==' 产品搜索!') {this.value='';}this.style.color='#fff';" onBlur="if(this.value=='') {this.value=' 产品搜索!';this.style.color='#fff';}" value=" 产品搜索!"   onKeyUp="keyupdeal(event);" onKeyDown="keydowndeal(event);" onClick="keyupdeal(event);"></td>
        <td width="40"><input type="image" src="Img/Ap.gif"></td>
    </tr>
  </Form>      
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="Img/cpgl.jpg" width="205" height="40"></td>
      </tr>    
      <tr>
        <td <%If ClassMenu = 1 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_ProductMain.asp"><img src="Img/cp1.png" align="absmiddle">　产品列表</a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr>  
            <tr>
        <td <%If ClassMenu = 2 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_ProductAdd.asp"><img src="Img/cp2.png" align="absmiddle">　添加产品</a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr>
            <tr>
        <td <%If ClassMenu = 3 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_ClassMain.asp"><img src="Img/cp3.png" align="absmiddle">　分类管理</a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr> 
                  <tr>
        <td <%If ClassMenu = 4 Then%>class="C_Menu1"<%Else%>class="C_Menu"<%End if%>><a href="E_BrandMain.asp"><img src="Img/cp4.png" align="absmiddle">　品牌管理</a></td>
      </tr>
      <tr>
        <td height="1" bgcolor="313235"></td>
      </tr>   
    </table>