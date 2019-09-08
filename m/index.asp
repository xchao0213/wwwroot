<!--#include file="../Include/conn.asp"-->
<!--#include file="Include/config.asp" -->
<!--#include file="Include/Sql.Asp" -->
<!--#include file="Include/zych_sql.asp"-->
<!DOCTYPE HTML>
<html>
<head>
<title><%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="css/style.css" media="screen" />
<link rel="stylesheet" type="text/css" href="css/flexslider.css" media="screen" />
<link rel="stylesheet" type="text/css" href="css/jquery.jslides.css" media="screen" />
<script type="text/javascript" src="js/jQuery.js"></script> 
<script type="text/javascript" src="js/iscroll.js"></script> 
</head>
<body>
<div id="wrapper">
<!--Header Starts-->
<!--#include file="Include/top.asp"-->  
<!--Header Ends-->
<!--Section Starts-->
<section id="main">
<!--Block Ends-->
<div class="wraper">
<div id="scroll_pic_view" class="scroll_pic_view" style="overflow: hidden;">
<div id="scroll_pic_view_div" style="width:3840px;">
<ul id="scroll_pic_view_ul">
<% set rsf=server.CreateObject("adodb.recordset")
rsf.open "select * from indexflash order by px_id asc",conn,1,1
if rsf.eof and rsf.bof then
response.Write("<li>暂无网站主页图片记录</li>")
end if
while not rsf.eof%>
<li><a onClick="return false;"><img width="100%" src="<%=rsf("img")%>"></a></li><%
rsf.movenext
wend
rsf.close
set rsf=nothing%>
</ul>
</div>
<div>
<ol id="scroll_pic_nav" class="scroll_pic_nav">
<script>
(function(d, $){
var scrollPicView = d.getElementById("scroll_pic_view"),
scrollPicViewDiv = d.getElementById("scroll_pic_view_div"),
lis = scrollPicViewDiv.querySelectorAll("li"),
w = scrollPicView.offsetWidth,
len = lis.length;
for(var i=0; i<len; i++){
lis[i].style.width = w+"px";
if(i == len-1){
    scrollPicViewDiv.style.width = w * len + "px";
}
}
var scroll_pic_view = new iScroll('scroll_pic_view', { 
snap: true,
momentum: false,
hScrollbar: false,
useTransition: true,
onScrollEnd: function() {
    $("#scroll_pic_nav li").removeClass("on").eq(this.currPageX).addClass("on");
    var  list=$("#scroll_pic_nav li");
    for(var k=0;k<list.length;k++){
        if(k<this.currPageX)
            $(list[k]).addClass("left");
        else
            $(list[k]).removeClass("left");
    }												
}
});

var nav_lis = new Array(lis.length);
//d.write('<li class="on"><span>1</span></li>');
d.write('<li class="on"></li>');
for(var i=1; i<nav_lis.length; i++){
//d.write("<li><span>"+(i+1)+"</span></li>");											
d.write("<li></li>");				
}
})(document, $);
</script>
</ol>
</div>
</div></div>
<!--Block Starts-->
<div class="block_module paper_bh_white">
<h2>关于我们</h2>
<div class="content_container">
<% set rs=server.createobject("adodb.recordset") 
exec="select * from about order by px_id asc  " 'ID=1为单页分类ID
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
response.Write("暂无记录...")
else
response.Write(""&stvalue(DelHtml(rs("body")),400)&"...")
abouturl="about/Show.asp?id="&rs("id")&""
end if
rs.close
set rs=nothing%>
  <p class="center" style="margin-top:10px;"><a href="<%=abouturl%>" class="grey_bt_big btn"><span class="app"></span>查看详细</a></p>
</div>
</div>
<!--Block Ends-->
<div class="block_module paper_bh_white">
<h2>工程项目</h2>
<div class="hotImg">
<%set pro=server.CreateObject("adodb.recordset")
pro.open "select top 6 * from [case] where tuijian=1 order by id desc",conn,1,1
if pro.eof and pro.bof then
response.Write("<div class=""onxx"">暂无推荐产品记录</div>")
end if
while not pro.eof
url="Case/Show.asp?id="&pro("id")&""%>
<a href="<%=url%>"><img src="<%=pro("img")%>" width="148" height="100" /><p><%=stvalue(pro("title"),22)%></p></a>
<%pro.movenext
wend
pro.close
set pro=nothing%>
</div>
</div>
<!--Block Starts-->
<div class="block_module paper_bh_white">
<h2>高屋动态</h2>
<div class="content_container">
<% set rsn=server.createobject("adodb.recordset") 
exec="select top 8 * from [news] order by id desc  " 'top 8代表每栏目显示5条新闻
rsn.open exec,conn,1,1 
if rsn.eof and rsn.bof then
response.Write("&nbsp;此分类暂无新闻 !")
end if
while not rsn.eof %>
  <p><a href="news/show.asp?id=<%= rsn("id") %>"><%= rsn("title") %></a></p>
<% rsn.movenext
wend
rsn.close
set rsn=nothing %>
  
</div>
</div>
<div class="block_module paper_bh_white">
<h2>领导关怀</h2>
<div id="accordion_menu">
<%
set rs=server.createobject("adodb.recordset") 
exec="select * from leader order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无服务项目！</div>"
end if
do while not rs.eof
url="Service/Show.asp?id="&rs("id")&""%>
<h5><%=rs("name")%><span class="arrow"></span></h5>
<div class="ac_content">
<p align="center"><a href="<%=url%>" title="<%=rs("name")%>"><img src="<%= rs("img") %>" alt="<%=rs("name")%>" width="220" height="120" onerror="this.src='<%=dir%>images/nologo.jpg'"/></a></p>
<p><%=rs("shuomin")%></p>
<p align="right"><a href="<%=url%>" title="<%=rs("name")%>">MOER</a></p>
</div>
<%rs.movenext
Loop
rs.close
Set rs=nothing%>
</div>
</div>
<!--Block Ends-->
</section>
<!--Section Ends-->
<!--Footer Starts-->
<!--#include file="Include/bottom.asp" -->
<!--Footer Ends-->
</div>
<!-- JQuery -->
<script src="js/jquery.min.js"></script>
<!-- JQueryUI -->
<script src="js/jquery-ui-1.8.16.custom.min.js"></script>
<!--Simple Carousel-->
<script src="js/jquery.flexslider.js"></script>
<!-- Script Controls -->
<script src="js/site.js"></script>
<script type="text/javascript">
/* Slideshow Control */		
	jQuery('.flexslider').flexslider();
</script>
</body>
</html>
