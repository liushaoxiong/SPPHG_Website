//�ı����λ�ñ��--------------------------------------------------------------
function changeAdminFlag(Content){
   var row=parent.parent.headFrame.document.all.Trans.rows[0];
   row.cells[3].innerHTML = Content ;
   return true;
}
//ɾ�����ַ���ڵ�--------------------------------------------------------------
function ConfirmDelSort(Result,ID)
//ɾ����Ʒ����ڵ�--------------------------------------------------------------
{
   if (confirm("��ȷʵҪɾ�����ࡢ���༰����������Ϣ��Ŀ��"))
   {
       window.location.href=Result+".asp?Action=Del&ID="+ID
   } 
}
//����ڵ�չ�����۵�(����)-------------------------------------------------------------
function AddToSort(imagePath){
  window.opener.LPform.LPattern.focus();								
  window.opener.document.LPform.LPattern.value=imagePath;
  window.opener=null;
  window.close();
}
//���ѡ�����------------------------------------------------------------------------
function OpenScript(url,width,height)
{
  var win = window.open(url,"SelectToSort",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
//����ڵ�չ�����۵�-------------------------------------------------------------------
function EndSortChange(a,b)
{
	if(eval(a).style.display=='')
	{
		eval(a).style.display='none';
		eval(b).className='SortEndFolderOpen';
	}
	else
	{
		eval(a).style.display='';
		eval(b).className='SortEndFolderClose';
	}
}
function SortChange(a,b)
{
	if(eval(a).style.display=='')
	{
		eval(a).style.display='none';
		eval(b).className='SortFolderOpen';
	}
	else
	{
		eval(a).style.display='';
		eval(b).className='SortFolderClose';
	}
}
//ͨ��ѡ��ɾ����Ŀ����ѡ-ȫѡ��--------------------------------------------------------
function CheckOthers(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      if (e.checked==false)
      {
	     e.checked = true;
      }
      else
      {
	     e.checked = false;
      }
   }
}

function CheckAll(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      e.checked = true;
   }
}
//�����Ŀɾ����ʾ------------------------------------------------------------
function ConfirmDel(message)
{
   if (confirm(message))
   {
      document.formDel.submit()
   }
}
//�����������ݱ༭��-----------------------------------------------------------
function OpenDialog(sURL, iWidth, iHeight)
{
   var oDialog = window.open(sURL, "_EditorDialog", "width=" + iWidth.toString() + ",height=" + iHeight.toString() + ",resizable=no,left=0,top=0,scrollbars=no,status=no,titlebar=no,toolbar=no,menubar=no,location=no");
   oDialog.focus();
}
//���������ַ�����Ч�ԣ�0-9��a-z,-,_��-------------------------------------------
function voidNum(argValue) 
{
   var flag1=false;
   var compStr="1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-";
   var length2=argValue.length;
   for (var iIndex=0;iIndex<length2;iIndex++)
   {
	   var temp1=compStr.indexOf(argValue.charAt(iIndex));
	   if(temp1==-1) 
	   {
	      flag1=false;
			break;							
	   }
	   else
	   { flag1=true; }
   }
   return flag1;
} 
//������Ա��¼------------------------------------------------------------------------------
function CheckAdminLogin()
{
   var check; 
   if (!voidNum(document.AdminLogin.LoginName.value))
   {
	  alert("����ȷ�������Ա���ƣ���0-9,a-z,-_������ϵ��ַ�������");
      document.AdminLogin.LoginName.focus();
	  return false;
	  exit;
   }    
   if (!voidNum(document.AdminLogin.LoginPassword.value))
   {
	  alert("���������Ա���롣");
	  document.AdminLogin.LoginPassword.focus();
	  return false;
	  exit;
   }
   if (!voidNum(document.AdminLogin.VerifyCode.value))
   {
      alert("����ȷ������֤�롣");
      document.AdminLogin.VerifyCode.focus();
	  return false;
	  exit;
   }
   return true;
}
//���༭����Ա------------------------------------------------------------------------------
function CheckAdminEdit()
{
   if (document.editAdminForm.AdminName.value.length<3 || document.editAdminForm.AdminName.value.length>10 )
   {
	  alert("����ȷ�����¼������0-9,a-z,-_�������3-10λ���ַ�������");
      document.editAdminForm.AdminName.focus();
	  return false;
	  exit;
   }	
   var check; 
   if (!voidNum(document.editAdminForm.AdminName.value))
   {
	  alert("����ȷ�����¼������0-9,a-z,-_�������3-10λ���ַ�������");
      document.editAdminForm.AdminName.focus();
	  return false;
	  exit;
   }
}
//���༭��Ա--------------------------------------------------------------------------------
function CheckMemEdit()
{
   if (document.editMemForm.MemName.value.length<3 || document.editMemForm.MemName.value.length>16 )
   {
	  alert("����ȷ�����¼������0-9,a-z,-_�������3-16λ���ַ�������");
      document.editMemForm.MemName.focus();
	  return false;
	  exit;
   }	
   var check; 
   if (!voidNum(document.editMemForm.MemName.value))
   {
	  alert("����ȷ�����¼������0-9,a-z,-_�������3-16λ���ַ�������");
      document.editMemForm.MemName.focus();
	  return false;
	  exit;
   }
}
//����Ա�˳���¼��ʾ--------------------------------------------------------------------------
function AdminOut()
{
   if (confirm("�����Ҫ�˳����������"))
   location.replace("CheckAdmin.asp?AdminAction=Out")
}
//��ת���ڼ�ҳ-------------------------------------------------------------------------------
function GoPage(Myself)
{
   window.location.href=Myself+"Page="+document.formDel.SkipPage.value;
}
//���ѡ��·����ID������·���������ı�·��--------------------------------------------------------
function AddSort(SortName,ID,Path)
{
	window.opener.editForm.SortName.focus();								
	window.opener.document.editForm.SortName.value=SortName;
	window.opener.document.editForm.SortID.value=ID;
	window.opener.document.editForm.SortPath.value=Path;
    window.opener=null;
    window.close();
}
function AddSort2(SortName,ID,Path)
{
	window.opener.editForm.SortName.focus();								
	window.opener.document.editForm.SortName2.value=SortName;
	window.opener.document.editForm.SortID2.value=ID;
	window.opener.document.editForm.SortPath2.value=Path;
    window.opener=null;
    window.close();
}


//ѡ����ʼ����-----------------------------------------------------------------
var DS_x,DS_y;

function dateSelector()  //����dateSelector��������ʵ��һ��������ʽ�����������
{
  var myDate=new Date();
  this.year=myDate.getFullYear();  //����year���ԣ���ݣ�Ĭ��ֵΪ��ǰϵͳ��ݡ�
  this.month=myDate.getMonth()+1;  //����month���ԣ��·ݣ�Ĭ��ֵΪ��ǰϵͳ�·ݡ�
  this.date=myDate.getDate();  //����date���ԣ��գ�Ĭ��ֵΪ��ǰϵͳ���ա�
  this.inputName='';  //����inputName���ԣ���������name��Ĭ��ֵΪ�ա�ע�⣺��ͬһҳ�г��ֶ����������򣬲������ظ���name��
  this.display=display;  //����display������������ʾ���������
}

function display()  //����dateSelector��display����������ʵ��һ��������ʽ������ѡ���
{
  var week=new Array('��','һ','��','��','��','��','��');

  document.write("<style type=text/css>");
  document.write("  .ds_font td,span  { font: normal 12px ����; color: #000000; }");
  document.write("  .ds_border  { border: 1px solid #000000; cursor: hand; background-color: #DDDDDD }");
  document.write("  .ds_border2  { border: 1px solid #000000; cursor: hand; background-color: #DDDDDD }");
  document.write("</style>");

  document.write("<input style='width:72px;text-align:left;' class='textfield' id='DS_"+this.inputName+"' name='"+this.inputName+"' value='"+this.year+"-"+this.month+"-"+this.date+"' title=˫���ɽ��б༩ ondblclick='this.readOnly=false;this.focus()' onblur='this.readOnly=true' readonly>");
  document.write("<button style='width:60px;height:18px;font-size:12px;margin:1px;border:1px solid #A4B3C8;background-color:#DFE7EF;' type=button onclick=this.nextSibling.style.display='block' onfocus=this.blur()>ѡ������</button>");

  document.write("<div style='position:absolute;display:none;text-align:center;width:0px;height:0px;overflow:visible' onselectstart='return false;'>");
  document.write("  <div style='position:absolute;left:-60px;top:20px;width:142px;height:165px;background-color:#F6F6F6;border:1px solid #245B7D;' class=ds_font>");
  document.write("    <table cellpadding=0 cellspacing=1 width=140 height=20 bgcolor=#CEDAE7 onmousedown='DS_x=event.x-parentNode.style.pixelLeft;DS_y=event.y-parentNode.style.pixelTop;setCapture();' onmouseup='releaseCapture();' onmousemove='dsMove(this.parentNode)' style='cursor:move;'>");
  document.write("      <tr align=center>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=subYear(this) title='��С���'>&lt;&lt;</td>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=subMonth(this) title='��С�·�'>&lt;</td>");
  document.write("        <td width=52%><b>"+this.year+"</b><b>��</b><b>"+this.month+"</b><b>��</b></td>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=addMonth(this) title='�����·�'>&gt;</td>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=addYear(this) title='�������'>&gt;&gt;</td>");
  document.write("      </tr>");
  document.write("    </table>");

  document.write("    <table cellpadding=0 cellspacing=0 width=140 height=20 onmousedown='DS_x=event.x-parentNode.style.pixelLeft;DS_y=event.y-parentNode.style.pixelTop;setCapture();' onmouseup='releaseCapture();' onmousemove='dsMove(this.parentNode)' style='cursor:move;'>");
  document.write("      <tr align=center>");
  for(i=0;i<7;i++)
	document.write("      <td>"+week[i]+"</td>");
  document.write("      </tr>");
  document.write("    </table>");

  document.write("    <table cellpadding=0 cellspacing=2 width=140 bgcolor=#EEEEEE>");
  for(i=0;i<6;i++)
  {
    document.write("    <tr align=center>");
	for(j=0;j<7;j++)
      document.write("    <td width=10% height=16 onmouseover=if(this.innerText!=''&&this.className!='ds_border2')this.className='ds_border' onmouseout=if(this.className!='ds_border2')this.className='' onclick=getValue(this,document.all('DS_"+this.inputName+"'))></td>");
    document.write("    </tr>");
  }
  document.write("    </table>");

  document.write("    <span style=cursor:hand onclick=this.parentNode.parentNode.style.display='none'>���رա�</span>");
  document.write("  </div>");
  document.write("</div>");

  dateShow(document.all("DS_"+this.inputName).nextSibling.nextSibling.childNodes[0].childNodes[2],this.year,this.month)
}

function subYear(obj)  //��С���
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  myObj[0].innerHTML=eval(myObj[0].innerHTML)-1;
  dateShow(obj.parentNode.parentNode.parentNode.nextSibling.nextSibling,eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function addYear(obj)  //�������
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  myObj[0].innerHTML=eval(myObj[0].innerHTML)+1;
  dateShow(obj.parentNode.parentNode.parentNode.nextSibling.nextSibling,eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function subMonth(obj)  //��С�·�
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  var month=eval(myObj[2].innerHTML)-1;
  if(month==0)
  {
    month=12;
    subYear(obj);
  }
  myObj[2].innerHTML=month;
  dateShow(obj.parentNode.parentNode.parentNode.nextSibling.nextSibling,eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function addMonth(obj)  //�����·�
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  var month=eval(myObj[2].innerHTML)+1;
  if(month==13)
  {
    month=1;
    addYear(obj);
  }
  myObj[2].innerHTML=month;
  dateShow(obj.parentNode.parentNode.parentNode.nextSibling.nextSibling,eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function dateShow(obj,year,month)  //��ʾ���·ݵ���
{
  var myDate=new Date(year,month-1,1);
  var today=new Date();
  var day=myDate.getDay();
  var selectDate=obj.parentNode.parentNode.previousSibling.previousSibling.value.split('-');
  var length;
  switch(month)
  {
    case 1:
    case 3:
    case 5:
    case 7:
    case 8:
    case 10:
    case 12:
      length=31;
      break;
    case 4:
    case 6:
    case 9:
    case 11:
      length=30;
      break;
    case 2:
      if((year%4==0)&&(year%100!=0)||(year%400==0))
        length=29;
      else
        length=28;
  }
  for(i=0;i<obj.cells.length;i++)
  {
    obj.cells[i].innerHTML='';
    obj.cells[i].style.color='';
    obj.cells[i].className='';
  }
  for(i=0;i<length;i++)
  {
    obj.cells[i+day].innerHTML=(i+1);
    if(year==today.getFullYear()&&(month-1)==today.getMonth()&&(i+1)==today.getDate())
      obj.cells[i+day].style.color='red';
    if(year==eval(selectDate[0])&&month==eval(selectDate[1])&&(i+1)==eval(selectDate[2]))
      obj.cells[i+day].className='ds_border2';
  }
}

function getValue(obj,inputObj)  //��ѡ������ڴ��������
{
  var myObj=inputObj.nextSibling.nextSibling.childNodes[0].childNodes[0].cells[2].childNodes;
  if(obj.innerHTML)
    inputObj.value=myObj[0].innerHTML+"-"+myObj[2].innerHTML+"-"+obj.innerHTML;
  inputObj.nextSibling.nextSibling.style.display='none';
  for(i=0;i<obj.parentNode.parentNode.parentNode.cells.length;i++)
    obj.parentNode.parentNode.parentNode.cells[i].className='';
  obj.className='ds_border2'
}

function dsMove(obj)  //ʵ�ֲ������
{
  if(event.button==1)
  {
    var X=obj.clientLeft;
    var Y=obj.clientTop;
    obj.style.pixelLeft=X+(event.x-DS_x);
    obj.style.pixelTop=Y+(event.y-DS_y);
  }
}

function UpSortPicture()  //���¸��޸ĵ�ͼƬ
{
  document.getElementById("SortPicture").src = document.FolderForm.Picture.value;
}

// �����֤��ͼ������--------------------------------------------------------------------		
function UpVerifyCode(){
  document.getElementById("PhotoSN").src = "../Include/VerifyCode.asp?t="+Math.random();
}
// �����ʾ��ɫͼ���--------------------------------------------------------------------		
function Getcolor(img_val){
	var obj = document.getElementById("colourPalette");
	ColorImg = img_val;
	if (obj){
	obj.style.left = getOffsetLeft(ColorImg) + "px";
	obj.style.top = (getOffsetTop(ColorImg) + ColorImg.offsetHeight) + "px";
	if (obj.style.visibility=="hidden")
	{
	obj.style.visibility="visible";
	}else {
	obj.style.visibility="hidden";
	}
	}
}
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}
function setColor(color){
	if (ColorImg){ColorImg.style.color = color;ColorImg.value = color;}
	document.getElementById("colourPalette").style.visibility="hidden";
}

function ShowDialog(url, width, height) {
	var arr = showModalDialog(url, window, "dialogWidth:" + width + "px;dialogHeight:" + height + "px;help:no;scroll:no;status:no");
}