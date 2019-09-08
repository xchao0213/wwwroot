<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
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
<link href="images/style.css"  rel="stylesheet"media="screen" />
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<script src="js/jquery.lazyload.min.js" type="text/javascript"></script>
<script src="js/blocksit.min.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div id="header"><ul class="yh"><%=zych_photo_class%></ul></div>
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Photo/"><%=zych_class(wurl)%></a></span></div>
<div id="wrapper">
<div id="container">
<%id=request.QueryString("id")
if  not isnumeric(id)  then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
if id="" then
exec="select * from [Photo] order by id desc"
else
exec="select * from [Photo] where ssfl="&id&" order by id desc"
end if
rs.open exec,conn,1,3
if rs.eof then
response.Write "<div class=""item"">暂无记录！</div>"
else
while not rs.eof
if html=0 then
url=""&dir&"Photo/"&showPhoto&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Photo/ShowPhoto.asp?id="&rs("id")&""
end if%>
<div class="grid yh"> 
<div class="imgholder"><a href="<%=url%>" target="_blank"><img class="lazy" src="images/pixel.gif" data-original="<%=rs("img")%>" alt="<%= rs("title") %>" width="200" /></a></div>
<strong><%=rs("title")%></strong>
<p><%=stvalue(DelHtml(rs("body")),120) %></p>
<div class="meta"><a href="<%=url%>" target="_blank">点击查看>>></a></div>
</div>    
<%rs.movenext
wend
end if
rs.close
set rs=nothing%>	
</div>
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
<script>
$(function(){
	$("img.lazy").lazyload({		
		load:function(){
			$('#container').BlocksIt({
				numOfCol:5,
				offsetX: 8,
				offsetY: 8
			});
		}
	});	
	$(window).scroll(function(){
			// 当滚动到最底部以上50像素时， 加载新内容
		if ($(document).height() - $(this).scrollTop() - $(this).height()<50){
			$('#container').append($("#test").html());		
			$('#container').BlocksIt({
				numOfCol:5,
				offsetX: 8,
				offsetY: 8
			});
			$("img.lazy").lazyload();
		}
	});
	
	//window resize
	var currentWidth = 1100;
	$(window).resize(function() {
		var winWidth = $(window).width();
		var conWidth;
		if(winWidth < 660) {
			conWidth = 440;
			col = 2
		} else if(winWidth < 880) {
			conWidth = 660;
			col = 3
		} else if(winWidth < 1100) {
			conWidth = 880;
			col = 4;
		} else {
			conWidth = 1100;
			col = 5;
		}
		
		if(conWidth != currentWidth) {
			currentWidth = conWidth;
			$('#container').width(conWidth);
			$('#container').BlocksIt({
				numOfCol: col,
				offsetX: 8,
				offsetY: 8
			});
		}
	});
});
</script>

</body>
</html>
