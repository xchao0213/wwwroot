<!--#include file="../Include/conn.asp" -->
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
Page=Request.QueryString("page")
if page="" then
page=1
elseif not IsNumeric(page) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
response.end
end if
page=int(page)
if id="" then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>"
response.end
elseif not IsNumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>"
response.end
end if

set rs=server.createobject("adodb.recordset") 
exec="select * from culture where id="&id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此文化！</div>"
response.End()
else
title=rs("title")
qx=rs("qx")
zz=rs("zz")
ly=rs("ly")
ssfl=rs("ssfl")
data=rs("data")
body=rs("body")
end if
rs.close
set rs=nothing
set fl=server.createobject("adodb.recordset") 
exec="select * from culture_fl where id="&ssfl
fl.open exec,conn,1,1
fl_id=fl("id")
fl_title=fl("title")
if html=0 then
newsdaohang="<a href="""&dir&"Culture/"&Newsfl&""&Separated&""&fl("id")&"."&HTMLName&""">"&fl("title")&"</a>"
elseif html=1 then
newsdaohang="<a href="""&dir&"Culture/Culturelist.asp?id="&fl("id")&""">"&fl("title")&"</a>"
end if
fl.close
set lf=nothing%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=title%>-<%=zych_class(wurl)%></title>
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="/">首页</a><a href="<%=dir%>Culture/"><%=zych_class(wurl)%></a><%=newsdaohang%></span></div>
<div class="w_650 fl">
<div class="main_content">
<h2 class="Showtitle"><%=title%></h2>
<div class="ShowDetails">作者：<%=zz%> 来源：<%=ly%> 发表时间：<%=data%>&nbsp;查看：<script src="/Include/Click.asp?click=Culture&id=<%=id%>"></script>次</div>
<div class="content-width" style="margin:10px; font-size:14px;">
<%'定义部分
if qx=9 and session("username")="" then
Response.Write("您没有权限浏览此内容,请登陆会员")
else
arr_Content=split(body,"[-------分----页-------码------]")
MaxPages=ubound(arr_Content)
Response.Write arr_Content(Page-1)
end if
%>

<%
if MaxPages >0 then
url=""&dir&"Culture/ShowCulture.asp?id="&id
	Response.write "<div class=""s_page""><a href='"& Url &"&page=1' title='第1页'>首页</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a  href='"& Url &"&page="& Prev_Page &"' title='第"& Prev_Page &"页'>上一页</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a  href='"& Url &"&page="& PageLink &"'>["& PageLink &"]</a> "
		else
			Response.Write "<span>["& PageLink &"]</span> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a href='" & Url & "&page=" & bdd_Page & "' title='第" & bdd_Page & "页'>下一页</A>"
	end if
	Response.write " <A href='" & Url & "&page=" & Maxpages+1 & "' title='第"& Maxpages+1 &"页'>尾页</A></div>"
end if
%>
 </div>
 <div class="fanye">
<%dim rstmp, nexttitle, prevtitle
set rstmp=server.CreateObject("adodb.recordset")
rstmp.open "select top 1 id, title from culture where id>"&id&" order by id asc",conn,1,1
if not rstmp.eof then
if html=0 then
nexttitle= "下一篇：<a href="""&dir&"Culture/"&ShowNews&Separated&rstmp(0)&"."&HTMLName&""">"&stvalue(rstmp(1),60)&"</a>"
elseif html=1 then
nexttitle="下一篇：<a  href="""&dir&"Culture/ShowCulture.asp?id="&rstmp(0)&""" title="""&stvalue(rstmp(1),60)&""">"&stvalue(rstmp(1),60)&"</a>"
end if
else
nexttitle = "下一篇：<a>已经没有了！</a>"
end if
rstmp.close
rstmp.open "select top 1 id, title from culture where id<"&id&" order by id desc"
if not rstmp.eof then
if html=0 then
prevtitle="上一篇：<a href="""&dir&"Culture/"&ShowNews&Separated&rstmp(0)&"."&HTMLName&""" >"&stvalue(rstmp(1),60)&"</a>"
elseif html=1 then
prevtitle="上一篇：<a href="""&dir&"Culture/ShowCulture.asp?id="&rstmp(0)& """ title="""&stvalue(rstmp(1),60)&""">"&stvalue(rstmp(1),60)&"</a>"
end if
else
prevtitle = "上一篇：<a>已经没有了！</a>"
end if
rstmp.close
set rstmp=nothing%>
<dd><%=prevtitle%></dd><dd><%=nexttitle%></dd></div>
</div>
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
<div class="clear"></div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
