<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if not isnumeric(id) then
response.write "<script>alert('警告!请勿尝试非法注入！');window.location.href='index.asp';</script>" 
Response.End()
end if
Page=Request.QueryString("page")
if page="" then
page=1
elseif not IsNumeric(page) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
response.end
end if
page=int(page)
set rsa=server.createobject("adodb.recordset")
if id<>"" then 
exec="select * from about where id="&id
else
exec="select * from about order by px_id asc"
end if
rsa.open exec,conn,1,1
if rsa.eof and rsa.bof then
abouttitle="无此标题"
else
aboutid=rsa("id")
abouttitle=rsa("title")
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=abouttitle%>_<%=zych_home%></title>
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
<div class="bar"><b><%=abouttitle%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>About/?id=<%=aboutid%>"><%=abouttitle%></a></span></div>
<div class="w_650 fl">
<div class="content-width">
<%if rsa.eof and rsa.bof then
response.Write "<div style=""padding:10px"">无此单页！</div>"
else%>
<%'定义部分
News_Content=""&rsa("body")&""
arr_Content=split(News_Content,"[-------分----页-------码------]")
MaxPages=ubound(arr_Content)
Response.Write arr_Content(Page-1)

if MaxPages >0 then
url=""&dir&"About/?id="&aboutid
	Response.write "<div class=""s_page""><a href='"& Url &"&page=1' title='第1页'>首页</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a  href='"& Url &"&page="& Prev_Page &"' title='第"& Prev_Page &"页'>上一页</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a  href='"&Url&"&page="&PageLink&"'>"&PageLink&"</a> "
		else
			Response.Write "<span>"&PageLink&"</span> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a href='" & Url & "&page="&bdd_Page& "' title='第" & bdd_Page & "页'>下一页</A>"
	end if
	Response.write " <A href='" & Url & "&page="&Maxpages+1& "' title='第"& Maxpages+1 &"页'>尾页</A></div>"
end if%>

<%end if
rsa.close
set rsa=nothing%></div>
</div>
<div class="w_280 fr">
<div class="box"><div class="title tubiao_A">企业导航 <span>About us</span></div>
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
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
