<!--#include file="../../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from [Products] where id="& id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此记录！</div>"
response.End()
end if
set BigClass=server.createobject("adodb.recordset") 
exec="select * from [bigClass] where BigClassID="&rs("BigClassID")&""
BigClass.open exec,conn,1,1
BigClassName=BigClass("BigClassName")

if rs("SmallClassID")<>0 then
set Smallclass=server.createobject("adodb.recordset") 
exec="select * from [smallclass] where SmallClassID="&rs("SmallClassID")&""
smallclass.open exec,conn,1,1
SmallClassName=Smallclass("SmallClassName")
end if%>
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=rs("title")%>_<%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="../css/style.css" media="screen" />
</head>
<body>
<div id="wrapper">
<!--Header Starts-->
<!--#include file="../Include/top.asp" -->
<!--Header Ends-->
<!--Section Starts-->
<section id="main">
<!--Block Starts-->
<div class="block_module paper_bh_white">
<div class="page blog"> <span class="post_date">所属分类：<%=BigClassName%></span>
<h1><a><%=rs("title")%></a></h1>
<div class="post_info"><img src="<%=Replace(rs("img"),"../",dir)%>" width="300"/></div>
<div class="content-width clear">
<%=rs("body")%>
</div>

<div class="tnex">
<%
 set rstmp=server.CreateObject("adodb.recordset")
 rstmp.open "select top 1 id, title from [Products] where id>"&request.QueryString("id")&" and BigClassID="&rs("BigClassID")&" order by id asc",conn,1,1
 if not rstmp.eof then
 nexttitle="<p><span>下一篇</span><a  href=""Show.asp?id="&rstmp(0)&""">"&stvalue(rstmp(1),34)&"</a></p>"
 else
 nexttitle = "<p><span>下一篇</span><a>已经没有了！</a></p>"
 end if
 rstmp.close
 rstmp.open "select top 1 id, title from [Products] where id<"&request.QueryString("id")&" and BigClassID="&rs("BigClassID")&" order by id desc"
 if not rstmp.eof then
 prevtitle="<p><span>上一篇</span><a href=""Show.asp?id="&rstmp(0)&""">"&stvalue(rstmp(1),34)&"</a></p>"
 else
 prevtitle ="<p><span>上一篇</span><a>已经没有了！</a></p>"
 end if
 rstmp.close
 set rstmp=nothing%>
<%=prevtitle%><%=nexttitle%> 
</div>
</div>
</div>
<!--Block Ends-->
<!--Block Starts-->
</section>
<!--Section Ends-->
<!--Footer Starts-->
<!--#include file="../Include/bottom.asp" -->
</div>
</body>
</html>
