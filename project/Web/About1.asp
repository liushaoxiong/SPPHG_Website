<!--#include file="../E_Inc/Conn.asp"-->
<!--#include file="../E_Inc/Function.asp"-->
<!--#include file="../E_Inc/E_Config.asp"-->
<!--#include file="../E_Inc/E_FunPassWord.asp"-->
<%
ClassID = int(request("ID"))
Set rs = Server.CreateObject("ADODB.Recordset")
if ClassID="" or not IsNumeric(ClassID) then
  sql="select top 1 * from E_About"
else
  sql="select * from E_About where ID="&ClassID
end if  
rs.open sql,conn,1,3

TiTle=rs("TiTle")	
Content=rs("Content")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="Keywords" content="<%=ClientkeyWords%>" />
<meta name="Description" content="<%=Clientdescription %>" />
<title><%=Title%>|<%=ClientTitle%>|<%=ClientkeyWords%></title>
<link href="../E_img/Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="../Web/Top.html" -->
<table width="984" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="202" valign="top"><!--#include file="../Web/Left.html" --></td>
    <td align="right" valign="top"  class="P2"><table width="98%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="30" height="30" align="left"><img src="../WebImages/Icon.jpg" width="21" height="20" /></td>
        <td align="left" class="T14"><strong><%=AboutName%></strong></td>
        <td align="right">����ǰ��λ�ã�<a href="/">��ҳ</a>&nbsp;&gt;&nbsp;<%=AboutName%></td>
      </tr>
    </table>
        <table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="4" bgcolor="e5e5e5"></td>
          </tr>
        </table>
        <table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="20"></td>
          </tr>
          <tr>
            <td align="left" class="T14 line25"><%=Content%></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
<%If ClassID = 5 Then%>
                    <tr>
            <td height="30">
<!--���ðٶȵ�ͼAPI-->
<style type="text/css">
    html,body{margin:0;padding:0;}
    .iw_poi_title {color:#CC5522;font-size:14px;font-weight:bold;overflow:hidden;padding-right:13px;white-space:nowrap}
    .iw_poi_content {font:12px arial,sans-serif;overflow:visible;padding-top:4px;white-space:-moz-pre-wrap;word-wrap:break-word}
</style>
<script type="text/javascript" src="http://api.map.baidu.com/api?key=&v=1.1&services=true"></script>

  <!--�ٶȵ�ͼ����-->
  <div style="width:730px;height:300px;border:#ccc solid 1px;" id="dituContent"></div>

<script type="text/javascript">
    //�����ͳ�ʼ����ͼ������
    function initMap(){
        createMap();//������ͼ
        setMapEvent();//���õ�ͼ�¼�
        addMapControl();//���ͼ��ӿؼ�
        addMarker();//���ͼ�����marker
    }
    
    //������ͼ������
    function createMap(){
        var map = new BMap.Map("dituContent");//�ڰٶȵ�ͼ�����д���һ����ͼ
        var point = new BMap.Point(103.724879,36.054296);//����һ�����ĵ�����
        map.centerAndZoom(point,16);//�趨��ͼ�����ĵ�����겢����ͼ��ʾ�ڵ�ͼ������
        window.map = map;//��map�����洢��ȫ��
    }
    
    //��ͼ�¼����ú�����
    function setMapEvent(){
        map.enableDragging();//���õ�ͼ��ק�¼���Ĭ������(�ɲ�д)
        map.enableScrollWheelZoom();//���õ�ͼ���ַŴ���С
        map.enableDoubleClickZoom();//�������˫���Ŵ�Ĭ������(�ɲ�д)
        map.enableKeyboard();//���ü����������Ҽ��ƶ���ͼ
    }
    
    //��ͼ�ؼ���Ӻ�����
    function addMapControl(){
        //���ͼ��������ſؼ�
	var ctrl_nav = new BMap.NavigationControl({anchor:BMAP_ANCHOR_TOP_LEFT,type:BMAP_NAVIGATION_CONTROL_LARGE});
	map.addControl(ctrl_nav);
                //���ͼ����ӱ����߿ؼ�
	var ctrl_sca = new BMap.ScaleControl({anchor:BMAP_ANCHOR_BOTTOM_LEFT});
	map.addControl(ctrl_sca);
    }
    
    //��ע������
    var markerArr = [{title:"�����򽭵�����е�豸�������ι�˾",content:"��ַ������ʡ����������������������ƺ�򽯼�ƺ��<br/>��ϵ�ˣ�л����<br/>�ֻ���13893315850<br/>�绰��0931-2881346<br/>���棺0931-2882996",point:"103.721089|36.054034",isOpen:0,icon:{w:21,h:21,l:0,t:0,x:6,lb:5}}
		 ];
    //����marker
    function addMarker(){
        for(var i=0;i<markerArr.length;i++){
            var json = markerArr[i];
            var p0 = json.point.split("|")[0];
            var p1 = json.point.split("|")[1];
            var point = new BMap.Point(p0,p1);
			var iconImg = createIcon(json.icon);
            var marker = new BMap.Marker(point,{icon:iconImg});
			var iw = createInfoWindow(i);
			var label = new BMap.Label(json.title,{"offset":new BMap.Size(json.icon.lb-json.icon.x+10,-20)});
			marker.setLabel(label);
            map.addOverlay(marker);
            label.setStyle({
                        borderColor:"#808080",
                        color:"#333",
                        cursor:"pointer"
            });
			
			(function(){
				var index = i;
				var _iw = createInfoWindow(i);
				var _marker = marker;
				_marker.addEventListener("click",function(){
				    this.openInfoWindow(_iw);
			    });
			    _iw.addEventListener("open",function(){
				    _marker.getLabel().hide();
			    })
			    _iw.addEventListener("close",function(){
				    _marker.getLabel().show();
			    })
				label.addEventListener("click",function(){
				    _marker.openInfoWindow(_iw);
			    })
				if(!!json.isOpen){
					label.hide();
					_marker.openInfoWindow(_iw);
				}
			})()
        }
    }
    //����InfoWindow
    function createInfoWindow(i){
        var json = markerArr[i];
        var iw = new BMap.InfoWindow("<b class='iw_poi_title' title='" + json.title + "'>" + json.title + "</b><div class='iw_poi_content'>"+json.content+"</div>");
        return iw;
    }
    //����һ��Icon
    function createIcon(json){
        var icon = new BMap.Icon("http://dev.baidu.com/wiki/static/map/API/img/ico-marker.gif", new BMap.Size(json.w,json.h),{imageOffset: new BMap.Size(-json.l,-json.t),infoWindowOffset:new BMap.Size(json.lb+5,1),offset:new BMap.Size(json.x,json.h)})
        return icon;
    }
    
    initMap();//�����ͳ�ʼ����ͼ
</script>
</td>
          </tr>
                    <tr>
            <td align="left" style="padding-top:10px"><img src="../WebImages/Map.jpg" width="737" height="308" /></td>
          </tr>
<%End if%>
          <tr>
            <td height="30">&nbsp;</td>
          </tr>
        </table></td>
  </tr>
</table>
<!--#include file="../Web/Foot.html" -->
</body>
</html>
<%rs.close
set rs=nothing
%>