<%
'取得大类和小类默认项
Dim strBClass,strSClass,strFormNameInfo,strFromCsSClass
strBClass=Int(Request.QueryString("BClass"))
strSClass=Int(Request.QueryString("SClass"))
'调用：Call InfoBClass2SClass(strBClass,strSClass,strFormNameInfo,strFromCsSClass)

'子例程开始
'======================================================================
Public Sub BClass2SClass(strBClass,strSClass,strFormNameInfo,strFromCsSClass,intClassType)

%>
	<script language="JavaScript">
	<!--
	SClasses = new Array(); 
	<%
	Dim z,bSqlss,InfoRs
	Set InfoRs = Server.CreateObject("ADODB.Recordset")
		bSqlss="select * from BClass Where IfShow = 0 order by ID asc "
		InfoRs.Open bSqlss, Conn, 1
			For z = 0 To InfoRs.Recordcount
			If InfoRs.EOF Then     
			Exit For 
			End If
		
				Response.Write("SClasses["&z&"] = new Array('"&Int(InfoRs("ID"))&"','")
				Response.Write(SClassId(Int(InfoRs("ID")),IntClassType))
				Response.Write("','")
				Response.Write(SClassContents(Int(InfoRs("ID")),IntClassType))
				Response.Write("');")
		InfoRs.MoveNext 
		next
	InfoRs.close
	set InfoRs = nothing
	%>
		
	function ChangeSClass(getBClass)
		{
		var getBClass = getBClass;
		var o,p,q;
		with(document.<%=strFormNameInfo%>)
			{
			SClass.length = 0 ; 
			SClass.options[SClass.length] = new Option("选择小类",""); 
			for (o = 0 ;o <SClasses.length;o++)
			  {
			  if(SClasses[o][0]==getBClass)
				{ 
				tmpSClasses = SClasses[o][1].split("|")
				tmpSClassesName = SClasses[o][2].split("|")
				for(p=0;p<tmpSClasses.length;p++)
				  {
				SClass.options[SClass.length] = new Option(tmpSClassesName[p],tmpSClasses[p]); 
				  }
				} 
			  } 
			for (q=0;q<SClass.options.length;q++)
			  { 
			  if(SClass.options[q].value=='<%=strSClass%>')
			  {SClass.options[q].selected=true;break;} 
			  }
			}
		}
		//-->
	</script>

	<select name="BClass" id="BClass" onChange = "ChangeSClass(this.options[this.selectedIndex].value)"  class=<%=strFromCsSClass%>>
		<option value="0">选择大类</option>
	<%
	Set InfoRs = Server.CreateObject("ADODB.Recordset")
		bSqlss="select * from BClass Where IfShow = 0 order by ID asc "
		InfoRs.Open bSqlss, Conn, 1
			For y = 0 To InfoRs.Recordcount
			If InfoRs.EOF Then     
			Exit For 
			End If
	%>		
		<option value=<%=InfoRs("ID")%>><%=InfoRs("BClassName")%></option>
	<%
		InfoRs.MoveNext 
		next
	InfoRs.close
	set InfoRs = nothing
	%>		
	</select>

	<select name="SClass" id="SClass"  class=<%=strFromCsSClass%>>
		<option value="">选择小类</option>
	</select>
	<script language="JavaScript">
		<!--
		with(document.<%=strFormNameInfo%>)
		{
		for(pp2=0;pp2<BClass.options.length;pp2++)
		  { 
		  if(BClass.options[pp2].value=='<%=strBClass%>')
		  {BClass.options[pp2].selected=true;break;} 
		  }
		ChangeSClass(BClass.options[BClass.selectedIndex].value)
		}
		//-->
	</script>
<%
End Sub
'======================================================================
'子例程结束


'[私有函数]取得数据库中小类表的数据（根据所属大类ID取小类的分类名称）
'===================================================================

Private  Function SClassContents(InfoRsId,iClassType)
Set SrSClass = Server.CreateObject("ADODB.Recordset")
	sSqlss="select * from SClass where IfShow = 0 And UpId = "&InfoRsId&" order by ID Asc"
	SrSClass.Open sSqlss, Conn, 1
	SClassContents=""
		Do while not SrSClass.eof
			SClassContents=SClassContents&SrSClass("SClassName")&"|"
			SrSClass.MoveNext
			Loop
if SClassContents<>"" then			
	LenNumInfo = len(SClassContents) - 1
	SClassContents = left(SClassContents,LenNumInfo)
	else
	SClassContents=""
	end if	
SrSClass.close
set SrSClass = nothing
End Function
'====================================================================

'[私有函数]取得数据库中小类表的数据（根据所属大类ID取小类的ID）
'===================================================================
Private  Function SClassId(InfoRsId2,iClassType2)
Set SrSClass2 = Server.CreateObject("ADODB.Recordset")
	sSqlss2="select * from SClass where  IfShow = 0 And UpId = "&InfoRsId2&" order by ID Asc"

	SrSClass2.Open sSqlss2, Conn, 1
	SClassId= ""
		Do while not SrSClass2.eof
			SClassId=SClassId&Int(SrSClass2("ID"))&"|"
			SrSClass2.MoveNext
			Loop
if SClassId<>"" then			
	LenNum2Info = len(SClassId) - 1
	SClassId = left(SClassId,LenNum2Info)
	else
	SClassId=""
	end if	
SrSClass2.close
set SrSClass2 = nothing
End Function
'====================================================================
%>

