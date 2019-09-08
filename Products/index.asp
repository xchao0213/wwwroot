<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"-->
<% 
BigClassID=request.QueryString("BigClassID")
SmallClassID=request.QueryString("SmallClassID")
if not isnumeric(BigClassID) or not isnumeric(SmallClassID) then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
if SmallClassID<>"" then 
set smallclass=server.createobject("adodb.recordset") 
exec="select * from [smallclass] where SmallClassID="&SmallClassID&""
smallclass.open exec,conn,1,1
bgname1=" > "&smallclass("SmallClassName")&""
bgname_1=" <a href="""&dir&"Products/?BigClassID="&smallclass("BigClassID")&"&SmallClassID="&smallclass("SmallClassID")&""" class=""on2"">"&smallclass("SmallClassName")&"</a>"
end if 
if BigClassID<>"" then 
set bigClass=server.createobject("adodb.recordset") 
exec="select * from [bigClass] where BigClassID="&BigClassID&""
bigClass.open exec,conn,1,1
bgname2=trim(bigClass("BigClassName"))
bgname_2=trim("<a href="""&dir&"Products/?BigClassID="&bigClass("BigClassID")&""">"&bigClass("BigClassName")&"</a>")
end if %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=zych_class(wurl)%>_<%=bgname2%><%=bgname1%>_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<script type="text/javascript" src="<%=dir%>js/jquery-1.4.2.min.js"></script>
<script type="text/javascript">
$(document).ready(function()
{
	$("#secondpane div.p_menu_1").mouseover(function()
    {
	     $(this).css({background:'#dfdfdf',color:'#ff0'}).next("div.p_menu_2").slideDown(300).siblings("div.p_menu_2").slideUp("slow");
         $(this).siblings().css({background:'#ededed',Color:'#FFF'});
	});
});
</script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="/">首页</a><a href="<%=dir%>Products/"><%=zych_class(wurl)%></a><%=bgname_2%><%=bgname_1%></span></div>
<div class="w_650 fl">
<div id="main_960">
<%if  not isnumeric(BigClassID)  or not isnumeric(SmallClassID) then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
if BigClassID="" then
exec="select  * from products order by px_id desc,id desc"
end if
if BigClassID<>"" then
exec="select  * from products where BigClassID="&BigClassID&" order by px_id desc,id desc"
end if
if BigClassID<>"" and SmallClassID<>"" then
exec="select  * from products where BigClassID="&BigClassID&" and SmallClassID="&SmallClassID&" order by px_id desc,id desc"
end if
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无产品</div>"
else
rs.PageSize =12
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
For i=1 To x
if html=0 then
url=""&Dir&"Products/"&ShowProducts&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Products/ShowProducts.asp?id="&rs("id")&""
end if%>
<div class="Pro_box">
<a href="<%=url%>" title="<%=rs("title")%>"><img src="<%=rs("img")%>" width="234" height="150" alt="<%=rs("title")%>" onerror="this.src='../images/nologo.jpg'"/><p><%= stvalue(rs("title"),21) %></p></a>
</div>
<%rs.movenext
next%>
</div>
<div class="page"><%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing
%></div>
  </div>
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_P">产品分类 <span>Products Nav</span></div>
<ul class="p_list" id="secondpane">
<%=zych_Products_class%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
