<!--#include file="../../Include/conn.asp"-->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"-->
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=zych_class(murl)%>_<%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="../css/style.css" media="screen" />
</head>
<body>
<div id="wrapper">
<!--Header Starts-->
<!--#include file="../Include/top.asp"-->  
<!--Header Ends-->
<!--Section Starts-->
<section id="main">
<!--Block Starts-->
<% set rs_fl=server.createobject("adodb.recordset") 
exec="select * from [culture_fl] order by px_id asc" 
rs_fl.open exec,conn,1,1 
if rs_fl.eof and rs_fl.bof then
response.Write("&nbsp;暂无任何分类 !")
end if
do while not rs_fl.eof
if html=0 then
ncurl=""&dir&"News/"&Newsfl&""&Separated&""&rs_fl("id")&"."&HTMLName&""
elseif html=1 then
ncurl=""&dir&"Culture/Culturelist.asp?id="&rs_fl("id")&""
end if%>
<div class="block_module paper_bh_white">
<div class="page blog">
<h1><a href="<%=ncurl%>"><%=rs_fl("title") %></a></h1>
<%set rs=server.createobject("adodb.recordset") 
exec="select top 10 * from [culture] where ssfl="&rs_fl("id")&" order by id desc  " 'top 8代表每栏目显示5条新闻
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
response.Write("&nbsp;此分类暂无内容 !")
end if
while not rs.eof
if IsNull(rs("url")) or trim(rs("url")&"")="" then
if html=0 then
url=""&dir&"News/"&ShowNews&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Culture/ShowCulture.asp?id="&rs("id")&""
end if
else
url=""&rs("url")&""
end if
if rs("img")="" then
img="../images/nopic.gif"
else
img=rs("img")
end if%>
<div class="post_info">
<a href="Show.asp?id=<%=rs("id")%>" class="title"><%=stvalue(rs("title"),30)%></a><a class="more"><%=right("0"&year(rs("data")),4)&"/"&right("0"&month(rs("data")),2)&"/"&right("0"&day(rs("data")),2)%></a></div>
<%rs.movenext
wend
rs.close
set rs=nothing%>
</div>
</div>
<%rs_fl.movenext
loop
rs_fl.close
set rs_fl=nothing%>
<!--Block Ends-->
</section>
<!--#include file="../Include/bottom.asp" -->
<!-- JQuery -->
<script src="js/jquery.min.js"></script>
<script src="js/jquery-ui-1.8.16.custom.min.js"></script>
<script src="js/slider.js"></script>
<script src="js/site.js"></script>
</body>
</html>
