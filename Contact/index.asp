<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=zych_class(wurl)%>_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<script type="text/javascript" src="http://api.map.baidu.com/api?key=&v=1.1&services=true"></script>
<style type="text/css">
.iw_poi_title {color:#CC5522;font-size:14px;font-weight:bold;overflow:hidden;padding-right:13px;white-space:nowrap}
.iw_poi_content {font:12px arial,sans-serif;overflow:visible;padding-top:4px;white-space:-moz-pre-wrap;word-wrap:break-word}
</style>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>Contact/"><%=zych_class(wurl)%></a></span></div>
<div class="Contact">
<div class="xinxi">
<div class="Name">
<p class="t1"><%=ditu_title%></p>
<p class="t2">ShangHai Gao Wu Zhi Ye CO.LTD</p></div>
<ul>
<li>��ϵ�绰��<%=zych_tel%></li>
<li>������룺<%=zych_fax%></li>
<li>E-MAIL ��<%=zych_mail%></li>
<li></li>
<li>�ٷ���վ��<%=zych_url%></li>
<li>��˾��ַ��<%=zych_dz%></li>
<li>�������룺<%=zych_Postal%></li>
</li></ul>
</div>
<div class="Message">
<div class="pdnav"><h1>�������� <span>Message</span></h1><p style="text-align:right"><a href="Message.asp" target="_blank">�鿴����</a></p></div>
<ul><form action="?action=bookadd" method="post" name="add" onSubmit="return check()">
<li><span class="t">�� ��</span><span class="i"><input name="title" type="text" class="input" size="40" maxlength="13"/></span></li>
<li><span class="t">�� ��</span><span class="i"><input name="name" type="text" class="input" size="40" maxlength="5"/></span></li>
<li><span class="t">�� ��</span><span class="i"><input name="tel" type="text" class="input" size="40" maxlength="11"/></span></li>
<li><span class="t">�� ��</span><span class="i"><input name="qq" type="text" class="input" size="40" maxlength="11"/>
<input name="sh" type="hidden" value="<%=zych_booksh%>" size="10"/></span></li>
<li><span class="t">�� ��</span><span class="i"><textarea name="body" cols="40" rows="3"  class="textarea"></textarea></span></li>
<li style="height:85px"><span class="t">�� ֤</span><span class="i"><input name="VerifyCode" class="v" type="text" size="8" maxlength="11"/><img src="../Include/safecode.asp" alt="������,����ˢ����֤��!" width="100" height="28" class="v" style="cursor:pointer; padding:0px 2px" onClick="this.src='../Include/safecode.asp?t='+(new Date().getTime());"> <span class="v" style="line-height:25px">��������֤��</span></span></li>
<li><span class="t">�� ��</span><span class="i"><button type="submit"  class="submit">�ύ����</button></span></li>
</form>
</ul>
</div>
</div>
<div class="bar"><b>��ͼλ��<i>Map Location</i></b></div>
<div style="width:990px;height:320px; margin:0 auto;border:#ccc solid 1px;" id="dituContent"></div>
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
        var point = new BMap.Point(<%=ditu_Coordinate%>);//����һ�����ĵ�����
        map.centerAndZoom(point,17);//�趨��ͼ�����ĵ�����겢����ͼ��ʾ�ڵ�ͼ������
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
        //���ͼ���������ͼ�ؼ�
	var ctrl_ove = new BMap.OverviewMapControl({anchor:BMAP_ANCHOR_BOTTOM_RIGHT,isOpen:1});
	map.addControl(ctrl_ove);
        //���ͼ����ӱ����߿ؼ�
	var ctrl_sca = new BMap.ScaleControl({anchor:BMAP_ANCHOR_BOTTOM_LEFT});
	map.addControl(ctrl_sca);
    }
    <%ditu_dinate= split(ditu_Coordinate,",")%>
    //��ע������
    var markerArr = [{title:"<%=ditu_mc%>",content:"<%=ditu_sm%><br/>��ַ:<%=zych_dz%>",point:"<%=ditu_dinate(0)%>|<%=ditu_dinate(1)%>",isOpen:1,icon:{w:23,h:25,l:46,t:21,x:9,lb:12}}
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
        var icon = new BMap.Icon("http://app.baidu.com/map/images/us_mk_icon.png", new BMap.Size(json.w,json.h),{imageOffset: new BMap.Size(-json.l,-json.t),infoWindowOffset:new BMap.Size(json.lb+5,1),offset:new BMap.Size(json.x,json.h)})
        return icon;
    }
    
    initMap();//�����ͳ�ʼ����ͼ
</script>

</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<%
if request("action")="bookadd" then
set rs=server.createobject("adodb.recordset")
sql="select * from ly"
rs.open sql,conn,1,3
title=request.form("title")
body=request.form("body")
name=request.form("name")
qq=request.form("qq")
mail=request.form("mail")
tel=request.form("tel")
sh=request.form("sh")
VerifyCode=request.form("VerifyCode") 
if title=""  then 
response.Write("<script language=javascript>alert('���Ա��ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if name=""  then 
response.Write("<script language=javascript>alert('��������Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if tel=""  then 
response.Write("<script language=javascript>alert('��ϵ�绰����Ϊ�գ���������ĵ绰ֻ�й���Ա�ܿ���!');history.go(-1)</script>") 
response.end 
end if
if qq=""  then 
response.Write("<script language=javascript>alert('��ϵQQ����Ϊ�գ����������QQֻ�й���Ա�ܿ���!');history.go(-1)</script>") 
response.end 
end if
IF sh="" or not isNumeric(request("sh")) then
response.write("<script>alert(""���棡�������ⲿ�ύ���ݣ�""); history.go(-1);</script>")
response.end
end if
if body=""  then 
response.Write("<script language=javascript>alert('���ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if  VerifyCode="" then 
response.Write("<script language=javascript>alert('��֤�벻��Ϊ��!');history.go(-1)</script>") 
response.end
end if 
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('��֤�����!');history.go(-1)</script>")
response.End
end if
rs.addnew
rs("title")=title
rs("body")=body
rs("name")=name
rs("qq")=qq
rs("mail")=mail
rs("tel")=tel
rs("sh")=sh
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('�����ύ�ɹ���');window.location.href='"&dir&"Contact/Message.asp';</script>" 
end if
%>