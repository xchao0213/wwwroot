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
page=int(page)
set rsa=server.createobject("adodb.recordset")
if id="" then
exec="select * from about order by id asc"
else
exec="select * from about where id="&id
end if
rsa.open exec,conn,1,1
if rsa.eof and rsa.bof then
abouttitle=""&zych_class(wurl)&""
else
abouttitle=rsa("title")
end if%>
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=abouttitle%>_<%=zych_home%></title>
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
<div class="page blog">
<h1><a href="#"><%=abouttitle%></a></h1>
<div class="content-width">
<%'定义部分
News_Content=""&rsa("body")&""
arr_Content=split(News_Content,"[-------分----页-------码------]")
MaxPages=ubound(arr_Content)
Response.Write arr_Content(Page-1)%>
<div class="Showpage">
<%url="Show.asp?id="&id
if MaxPages >0 then
	Response.write "<a href='"& Url &"&page=1' title='第1页'>首页</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a  href='"& Url &"&page="& Prev_Page &"' title='第"& Prev_Page &"页'>上一页</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a  href='"& Url &"&page="& PageLink &"'>["& PageLink &"]</a> "
		else
			Response.Write "<font color='#EF867B'><b>["& PageLink &"]</b></font> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a href='" & Url & "&page=" & bdd_Page & "' title='第" & bdd_Page & "页'>下一页</A>"
	end if
	Response.write " <A href='" & Url & "&page=" & Maxpages+1 & "' title='第"& Maxpages+1 &"页'>尾页</A>"
end if
%></div>
 </div>

    <div class="tags"> <span>相关页面:</span>
        <ul>
        <%set rsa=server.CreateObject("adodb.recordset")
rsa.open "select * from about where xianshi=1 order by px_id asc",conn,1,1
if rsa.eof and rsa.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not rsa.eof %>
<li><a href="Show.asp?id=<%=rsa("id")%>"><%=rsa("title")%></a></li>
<%rsa.movenext
wend
rsa.close
set rsa=nothing%>
        </ul>
    </div>
</div>
</div>
<!--Block Ends-->
<!--Block Starts-->
</section>
<!--Section Ends-->
<!--Footer Starts-->
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
