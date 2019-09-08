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
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>News/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<div class="list">
<% 
set news_fl=server.createobject("adodb.recordset") 
exec="select * from [news_fl] order by px_id asc  " 
news_fl.open exec,conn,1,1 
if news_fl.eof and news_fl.bof then
response.Write("&nbsp;暂无任何分类 !")
end if
do while not news_fl.eof
if html=0 then
ncurl=""&dir&"News/"&Newsfl&""&Separated&""&news_fl("id")&"."&HTMLName&""
elseif html=1 then
ncurl=""&dir&"News/Newslist.asp?id="&news_fl("id")&""
end if%>
<div class="dt"><div class="N1">NO</div><div class="N2"><%=news_fl("title")%></div><div class="N3"><a href="<%= ncurl %>" style="color:#FFF">MORE</a></div></div>
<%set news=server.createobject("adodb.recordset") 
exec="select top 8 * from [news] where ssfl="&news_fl("id")&" order by id desc  " 'top 8代表每栏目显示5条新闻
news.open exec,conn,1,1 
if news.eof and news.bof then
response.Write("&nbsp;此分类暂无新闻 !")
end if
while not news.eof
i=(i+1)
data=right("0"&year(news("data")),4)&"-"&right("0"&month(news("data")),2)&"-"&right("0"&day(news("data")),2)
if IsNull(news("url")) or trim(news("url")&"")="" then
if html=0 then
url=""&dir&"News/"&ShowNews&""&Separated&""&news("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"News/ShowNews.asp?id="&news("id")&""
end if
else
url=""&news("url")&""
end if
if news("img")<>"" then img=news("img") else img="../images/nologo.jpg"
If newslb<>1 Then %>
<div class="dl"><div class="N1"><%=right("00"&i,3)%></div>
<div class="N2"><a href="<%=url%>"><img src="<%=img%>" width="110" height="70" alt="<%=stvalue(news("title"),60)%>" onerror="this.src='../images/nologo.jpg'" /></a>
<a href="<%=url%>" class="t" title="<%=stvalue(news("title"),60)%>"><%=stvalue(news("title"),60)%></a><p><%=stvalue(Delhtml(news("body")),100)%></p></div>
<div class="N3"><%=news("data")%></div></div>
<% Else %>
<div class="dd"><div class="N1"><%=right("00"&i,3)%></div><div class="N2"><a title="<%=news("title")%>" target="_blank" href="<%=url%>"><%=stvalue(news("title"),60)%></a></div><div class="N3"><%=news("data")%></div></div>
<%end if
news.movenext
wend
news.close
set news=nothing
news_fl.movenext
loop
news_fl.close
set news_fl=nothing%>
</div>
<div class="page"></div>
</div>
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_A">企业导航 <span>About us</span></div>
<ul class="list_2">
    <li><a href="/" title="首页" class="selected">首页</a></li>
    <li><a href="/About/" title="走进高屋" class="">走进高屋</a></li>
    <li><a href="/News/" title="高屋动态" class="">高屋动态</a></li>
    <li><a href="/about/?id=10" title="工程项目">工程项目</a></li>
    <li><a href="/Culture/" title="高屋文化" class="">高屋文化</a></li>
    <li><a href="/Fengfa/" title="奉发园区" class="">奉发园区</a></li>
    <li><a href="/Contact/" title="联系高屋" class="">联系高屋</a></li>  
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
