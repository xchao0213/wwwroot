<!--#include file="Include/conn.asp"-->
<!--#include file="Include/config.asp" -->
<!--#include file="Include/Sql.Asp" -->
<!--#include file="Include/zych_sql.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=zych_home%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<script type="text/javascript" src="js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="js/superslide.2.1.js"></script>
<script type="text/javascript" src="js/dropdown.js"></script>
</head>
<body>
<!--#include file="Include/top.asp" -->
<!-- 代码 开始 -->
<div class="fullSlide">
<div class="bd">
<ul><%=zych_index_flash%></ul>
</div>
<div class="hd">
<ul></ul>
</div>
<span class="aprev"></span> <span class="anext"></span> </div>

<script type="text/javascript">
jQuery(".fullSlide").hover(function() {
    jQuery(this).find(".aprev,.anext").stop(true, true).fadeTo("show", 0.5)
},
function() {
    jQuery(this).find(".aprev,.anext").fadeOut()
});
jQuery(".fullSlide").slide({
    titCell: ".hd ul",
    mainCell: ".bd ul",
    effect: "fold",
    autoPlay: true,
    autoPage: true,
    trigger: "click",
    startFun: function(i) {
        var curLi = jQuery(".fullSlide .bd li").eq(i);
        if ( !! curLi.attr("src")) {
            curLi.css("background-image", curLi.attr("src")).removeAttr("src")
        }
    }
});

</script>
<!-- 代码 结束 -->    
<div class="main">

<% set rsfl=server.CreateObject("adodb.recordset")
rsfl.open "select * from news_fl order by px_id asc",conn,1,1
if rsfl.eof and rsfl.bof then
response.Write("暂无记录")
else
' shu=rsfl.RecordCount
shu = 2
end if 
set rsf2=server.CreateObject("adodb.recordset")
rsf2.open "select top 1 * from culture_fl order by px_id asc",conn,1,1
if rsf2.eof and rsf2.bof then
response.Write("暂无记录")
else
shu2=rsf2.RecordCount
end if %>
<script type="text/javascript" language="javascript">
function g(o){return document.getElementById(o);}
function HoverLi(n){
for(var i=1;i<=<%=shu%>;i++){g('tb_'+i).className='normaltab';g('tbc_'+i).className='undis';}g('tbc_'+n).className='dis';g('tb_'+n).className='hovertab';
if(n==1){
g('more').href ="/News/"
}
else{
g('more').href="/Culture/"
}
}

</script>
<div class="in_news"><div class="title">
<div id="tb" class="tb"><ul>
<%while not rsfl.eof
if rsfl("px_id")=1 then hover="hovertab" else hover="normaltab"
Response.Write("<li id=""tb_"&rsfl("px_id")&""" class="""&hover&""" onMouseOver=""i:HoverLi("&rsfl("px_id")&");"">"&rsfl("title")&"</li>"&vbcrlf)
rsfl.movenext
wend
rsfl.close
set rsfl=nothing
while not rsf2.eof
hover="normaltab"
Response.Write("<li id=""tb_"&rsf2("px_id")&""" class="""&hover&""" onMouseOver=""i:HoverLi("&rsf2("px_id")&");"">"&rsf2("title")&"</li>"&vbcrlf)
rsf2.movenext
wend
rsf2.close
set rsf2=nothing %> 
</ul></div><a id='more' class="more" href="/News/">+MORE</a></div>
<div class="ctt">
<%set rsfl=server.CreateObject("adodb.recordset")
rsfl.open "select * from news_fl order by px_id asc",conn,1,1
if rsfl.eof and rsfl.bof then
response.Write("暂无记录")
else
shu=rsfl.RecordCount
end if
while not rsfl.eof
if rsfl("px_id")=1 then undis="dis" else undis="undis" 
%>
<ul class="<%=undis%>" id="tbc_<%=rsfl("px_id")%>">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%Set rs1=Server.CreateObject("ADODB.Recordset")
sql = "select top 8 * from news where ssfl="&rsfl("id")&" order by id desc"
rs1.open sql,conn,1,1
do while not rs1.eof
if IsNull(rs1("url")) or trim(rs1("url")&"")="" then
  if html<>0 then
  url=""&dir&"News/ShowNews.asp?id="&rs1("id")&""
  else
  url=""&dir&"News/"&ShowNews&Separated&rs1("id")&"."&HTMLName&""
  end if
else
  url=""&rs1("url")&""
end if%>
<tr><td width="77%"><a href="<%=url%>" class="left"><%=left(rs1("title"),20)%></a></td><td width="23%" align="right"><i><%=formatDate(rs1("Data"),1)%></i></td></tr>
<%rs1.movenext
loop 
rs1.close
set rs1=nothing%>
</table></ul>
<%rsfl.movenext
wend
rsfl.close
set rsfl=nothing %>

<%set rsf2=server.CreateObject("adodb.recordset")
rsf2.open "select top 1 * from culture_fl order by px_id asc",conn,1,1
if rsf2.eof and rsf2.bof then
response.Write("暂无记录")
else
shu2=rsf2.RecordCount
end if
while not rsf2.eof
if rsf2("px_id")=1 then undis="undis" else undis="undis" 
%>
<ul class="undis" id="tbc_<%=rsf2("px_id")%>">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%Set rs2=Server.CreateObject("ADODB.Recordset")
sql = "select top 8 * from culture where ssfl="&rsf2("id")&" order by id desc"
rs2.open sql,conn,1,1
do while not rs2.eof
if IsNull(rs2("url")) or trim(rs2("url")&"")="" then
  if html<>0 then
  url=""&dir&"Culture/ShowCulture.asp?id="&rs2("id")&""
  else
  url=""&dir&"News/"&ShowNews&Separated&rs2("id")&"."&HTMLName&""
  end if
else
  url=""&rs2("url")&""
end if%>
<tr><td width="77%"><a href="<%=url%>" class="left"><%=left(rs2("title"),20)%></a></td><td width="23%" align="right"><i><%=formatDate(rs2("Data"),1)%></i></td></tr>
<%rs2.movenext
loop 
rs2.close
set rs2=nothing%>
</table></ul>
<%rsf2.movenext
wend
rsf2.close
set rsf2=nothing%>
</div>
</div>

<div class="in_about"><div class="in_title"><b>公告<span>Notice</span></b><a class="more" href="/About/">+MORE</a></div>
<div class="box"><p><%set ra=server.createobject("adodb.recordset") 
exec="select * from about where id=7 order by id asc" 'ID=1为单页分类ID
ra.open exec,conn,1,1 
if ra.eof and ra.bof then
response.Write("暂无记录")
else
response.Write(""&stvalue(DelHtml(ra("body")),394)&"<a href=""/About"" target=""_blank"" ><b>[详情]</b></a>")
end if
ra.close
set ra=nothing%></p></div>
</div>

<div class="in_video"><div class="in_title"><b>视频展示<span>Video</span></b><a class="more" href="/Video/">+MORE</a></div>
<div class="video"><%=zych_index_video%></div>
</div>
</div>
<div class="main">
<%if Scroll=1 then table="Case" else table="Products"%>
<div class="in_al"><div class="in_title"><b><%=in_title%><span><%=table%></span></b><a class="more" href="/<%=table%>/">+MORE</a></div>
<div class="carousel-box">
<div class="inner">
<div class="carousel">
<ul>
<%shu=4*in_hs '注：3为行
total=shu*4 '注：4为共有多少屏
set rs=server.CreateObject("adodb.recordset")
rs.open "select top "&total&" * from ["&table&"] where tuijian=1 order by id desc",conn,1,1
if rs.eof and rs.bof then
response.Write("<a>暂无记录</a>")
end if
Response.Write("<li>")
while not rs.eof
i=i+1
if Scroll=1 then
	if html<>1 then
	url=""&Dir&"Case/"&ShowCase&Separated&rs("id")&"."&HTMLName&""
	else
	url=""&Dir&"Case/ShowCase.asp?id="&rs("id")&""
	end if
else
	if html<>1 then
	url=""&Dir&"Products/"&ShowProducts&Separated&rs("id")&"."&HTMLName&""
	else
	url=""&Dir&"Products/ShowProducts.asp?id="&rs("id")&""
	end if
end if
str=str&"<a href="""&url&""" target=""_blank""><img src="""&imgurl(rs("img"))&""" alt="""&rs("title")&""" width=""280"" height=""70"" /><p>"&stvalue(rs("title"),32)&"</p></a>"
if i mod shu=0 then
str=str&"</li><li>"
end if 
rs.movenext
wend
str=str&"</li>"
if right(str,9)="<li></li>" then
str=left(str,len(str)-9)
end if
Response.Write(str)
rs.close
set rs=nothing%>
</ul>
<button class="prev"></button>
<button class="next"></button>
</div>
</div>
</div>

</div>

</div>
<div class="main">
<div class="in_co"><div class="in_title"><b>联系我们<span>Contact</span></b><a class="more" href="/Contact/">+MORE</a></div>
<div class="conimg"></div>
<p>联系电话：<%=zych_tel%></p>
<p>联系邮箱：<%=zych_mail%></p>
<p>官方网站：<%=zych_url%></p>
<p>联系地址：<%=zych_dz%></p>
</div>
<div class="in_fw"><div class="in_title"><b>专题特刊<span>Special</span></b><a class="more" href="/Service/">+MORE</a></div>
<%set rs=server.createobject("adodb.recordset") 
exec="select top 1 * from project order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无服务项目！</div>"
end if
do while not rs.eof
if html=0 then
url=""&dir&"Service/"&Showproject&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Service/ShowProject.asp?id="&rs("id")&""
end if%>
<a href="<%=url%>" class="box"><dl><img src="<%=imgurl(rs("img"))%>" alt="<%=rs("name")%>" width="147px" height="90px" />
<dt><%=rs("name")%></dt>
</dl></a>
<%rs.movenext
Loop
rs.close
Set rs=nothing%>
</div>
<div class="in_fw"><div class="in_title"><b>领导关怀<span>Leader</span></b><a class="more" href="/Leader/">+MORE</a></div>
<%set rs=server.createobject("adodb.recordset") 
exec="select top 1 * from leader order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无服务项目！</div>"
end if
do while not rs.eof
if html=0 then
url=""&dir&"Service/"&Showproject&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Leader/ShowLeader.asp?id="&rs("id")&""
end if%>
<a href="<%=url%>" class="box"><dl><img src="<%=imgurl(rs("img"))%>" alt="<%=rs("name")%>" width="147px" height="90px" />
<dt><%=rs("name")%></dt>
</dl></a>
<%rs.movenext
Loop
rs.close
Set rs=nothing%>
</div>

<div id="in_link"><div class="t">友情链接</div><div class="line"></div>
<div class="box">
<p><%=zych_link_img%></p>
<p style="display:none"><%=zych_link_text%></p>
</div></div>
</div>

<!--#include file="Include/bottom.asp" -->
</body>
</html>
