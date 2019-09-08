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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Contact/"><%=zych_class(wurl)%></a></span></div>
<div class="Contact">
<div class="xinxi">
<div class="Name">
<p class="t1"><%=ditu_title%></p>
<p class="t2">ShangHai Gao Wu Zhi Ye CO.LTD</p></div>
<ul>
<li>联系电话：<%=zych_tel%></li>
<li>传真号码：<%=zych_fax%></li>
<li>E-MAIL ：<%=zych_mail%></li>
<li></li>
<li>官方网站：<%=zych_url%></li>
<li>公司地址：<%=zych_dz%></li>
<li>邮政编码：<%=zych_Postal%></li>
</li></ul>
</div>
<div class="Message">
<div class="pdnav"><h1>在线留言 <span>Message</span></h1><p style="text-align:right"><a href="Message.asp" target="_blank">查看留言</a></p></div>
<ul><form action="?action=bookadd" method="post" name="add" onSubmit="return check()">
<li><span class="t">标 题</span><span class="i"><input name="title" type="text" class="input" size="40" maxlength="13"/></span></li>
<li><span class="t">姓 名</span><span class="i"><input name="name" type="text" class="input" size="40" maxlength="5"/></span></li>
<li><span class="t">电 话</span><span class="i"><input name="tel" type="text" class="input" size="40" maxlength="11"/></span></li>
<li><span class="t">Ｑ Ｑ</span><span class="i"><input name="qq" type="text" class="input" size="40" maxlength="11"/>
<input name="sh" type="hidden" value="<%=zych_booksh%>" size="10"/></span></li>
<li><span class="t">内 容</span><span class="i"><textarea name="body" cols="40" rows="3"  class="textarea"></textarea></span></li>
<li style="height:85px"><span class="t">验 证</span><span class="i"><input name="VerifyCode" class="v" type="text" size="8" maxlength="11"/><img src="../Include/safecode.asp" alt="看不清,请点击刷新验证码!" width="100" height="28" class="v" style="cursor:pointer; padding:0px 2px" onClick="this.src='../Include/safecode.asp?t='+(new Date().getTime());"> <span class="v" style="line-height:25px">请填入验证码</span></span></li>
<li><span class="t">提 交</span><span class="i"><button type="submit"  class="submit">提交留言</button></span></li>
</form>
</ul>
</div>
</div>
<div class="bar"><b>地图位置<i>Map Location</i></b></div>
<div style="width:990px;height:320px; margin:0 auto;border:#ccc solid 1px;" id="dituContent"></div>
<script type="text/javascript">
    //创建和初始化地图函数：
    function initMap(){
        createMap();//创建地图
        setMapEvent();//设置地图事件
        addMapControl();//向地图添加控件
        addMarker();//向地图中添加marker
    }
    
    //创建地图函数：
    function createMap(){
        var map = new BMap.Map("dituContent");//在百度地图容器中创建一个地图
        var point = new BMap.Point(<%=ditu_Coordinate%>);//定义一个中心点坐标
        map.centerAndZoom(point,17);//设定地图的中心点和坐标并将地图显示在地图容器中
        window.map = map;//将map变量存储在全局
    }
    
    //地图事件设置函数：
    function setMapEvent(){
        map.enableDragging();//启用地图拖拽事件，默认启用(可不写)
        map.enableScrollWheelZoom();//启用地图滚轮放大缩小
        map.enableDoubleClickZoom();//启用鼠标双击放大，默认启用(可不写)
        map.enableKeyboard();//启用键盘上下左右键移动地图
    }
    
    //地图控件添加函数：
    function addMapControl(){
        //向地图中添加缩放控件
	var ctrl_nav = new BMap.NavigationControl({anchor:BMAP_ANCHOR_TOP_LEFT,type:BMAP_NAVIGATION_CONTROL_LARGE});
	map.addControl(ctrl_nav);
        //向地图中添加缩略图控件
	var ctrl_ove = new BMap.OverviewMapControl({anchor:BMAP_ANCHOR_BOTTOM_RIGHT,isOpen:1});
	map.addControl(ctrl_ove);
        //向地图中添加比例尺控件
	var ctrl_sca = new BMap.ScaleControl({anchor:BMAP_ANCHOR_BOTTOM_LEFT});
	map.addControl(ctrl_sca);
    }
    <%ditu_dinate= split(ditu_Coordinate,",")%>
    //标注点数组
    var markerArr = [{title:"<%=ditu_mc%>",content:"<%=ditu_sm%><br/>地址:<%=zych_dz%>",point:"<%=ditu_dinate(0)%>|<%=ditu_dinate(1)%>",isOpen:1,icon:{w:23,h:25,l:46,t:21,x:9,lb:12}}
		 ];
    //创建marker
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
    //创建InfoWindow
    function createInfoWindow(i){
        var json = markerArr[i];
        var iw = new BMap.InfoWindow("<b class='iw_poi_title' title='" + json.title + "'>" + json.title + "</b><div class='iw_poi_content'>"+json.content+"</div>");
        return iw;
    }
    //创建一个Icon
    function createIcon(json){
        var icon = new BMap.Icon("http://app.baidu.com/map/images/us_mk_icon.png", new BMap.Size(json.w,json.h),{imageOffset: new BMap.Size(-json.l,-json.t),infoWindowOffset:new BMap.Size(json.lb+5,1),offset:new BMap.Size(json.x,json.h)})
        return icon;
    }
    
    initMap();//创建和初始化地图
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
response.Write("<script language=javascript>alert('留言标题不能为空!');history.go(-1)</script>") 
response.end 
end if
if name=""  then 
response.Write("<script language=javascript>alert('姓名不能为空!');history.go(-1)</script>") 
response.end 
end if
if tel=""  then 
response.Write("<script language=javascript>alert('联系电话不能为空，请放心您的电话只有管理员能看到!');history.go(-1)</script>") 
response.end 
end if
if qq=""  then 
response.Write("<script language=javascript>alert('联系QQ不能为空，请放心您的QQ只有管理员能看到!');history.go(-1)</script>") 
response.end 
end if
IF sh="" or not isNumeric(request("sh")) then
response.write("<script>alert(""警告！请勿尝试外部提交数据！""); history.go(-1);</script>")
response.end
end if
if body=""  then 
response.Write("<script language=javascript>alert('内容不能为空!');history.go(-1)</script>") 
response.end 
end if
if  VerifyCode="" then 
response.Write("<script language=javascript>alert('验证码不能为空!');history.go(-1)</script>") 
response.end
end if 
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('验证码错误!');history.go(-1)</script>")
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
Response.Write "<script>alert('留言提交成功！');window.location.href='"&dir&"Contact/Message.asp';</script>" 
end if
%>