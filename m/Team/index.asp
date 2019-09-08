<!--#include file="../../Include/conn.asp"-->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/Pager.asp"-->
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
<div class="block_module paper_bh_white">
<div class="page blog">
<h1><%=zych_class(murl)%></h1>
<%
set rs=server.createobject("adodb.recordset") 
exec="select * from [tuandui] order by px_id asc"
rs.open exec,conn,1,3
if rs.eof then
response.Write "<div style=""padding:10px"">暂无记录！</div>"
else
rs.PageSize =6
iCount=rs.RecordCount 
iPageSize=rs.PageSize
maxpage=rs.PageCount 
page=request("page")
if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
if page=maxpage then
x=iCount-(maxpage-1)*iPageSize
else
x=iPageSize
end if
end if
For i=1 To x %>
<div class="team">
<img src="<%=Replace(rs("img"),"../",dir)%>" width="150" alt="<%=rs("name")%>" onerror="this.src='<%=dir%>images/nologo.jpg'"/>
<div class="tbox">
<p>姓名：<a><%=rs("name")%></a></p>
<p>职务：<samp><%=rs("zw")%></samp></p>
<p>专项：<samp><%=rs("unit")%></samp></p>
<p>博客：<samp><%=rs("url")%>【<a href="<%=rs("url")%>" target="_blank">访问</a>】</samp></p>
<p><%=stvalue(DelHtml(rs("body")),200) %></p></div>
</div>
<%rs.movenext
next%>
</div>
</div>
<div class="block_module paper_bh_white">
<div class="list_page"><%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%></div>
</div>
<!--Block Ends-->
</section>

<!--#include file="../Include/bottom.asp" -->
</div>
<!-- JQuery -->
<script src="js/jquery.min.js"></script>
<script src="js/jquery-ui-1.8.16.custom.min.js"></script>
<script src="js/slider.js"></script>
<script src="js/site.js"></script>
</body>
</html>
