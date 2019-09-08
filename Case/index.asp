<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%id=request.QueryString("id")
set fl=server.createobject("adodb.recordset") 
if id<>"" then
exec="select * from case_class where id="&id
else
exec="select * from case_class"
end if
fl.open exec,conn,1,1
fl_id=fl("id")
fl_title=fl("title")
if html=0 then
classtitle="<a href="""&dir&"Case/"&Casefl&""&Separated&""&fl("id")&"."&HTMLName&""">"&fl("title")&"</a>"
elseif html=1 then
classtitle="<a href="""&dir&"Case/?id="&fl("id")&""">"&fl("title")&"</a>"
end if
fl.close
set lf=nothing %>
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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Case/"><%=zych_class(wurl)%></a><%=classtitle %></span></div>
<div class="w_650 fl">
<div id="main_960"><%
if  not isnumeric(id)  then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
if id="" then
exec="select * from [case] order by id desc"
else
exec="select * from [case] where ssfl="&id&" order by id desc"
end if
rs.open exec,conn,1,3
if rs.eof then
response.Write "<div style=""padding:10px"">暂无记录！</div>"
else
rs.PageSize =10
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
url=""&dir&"Case/"&showcase&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Case/ShowCase.asp?id="&rs("id")&""
end if%>
<div class="case_box">
<a href="<%=url%>" title="<%=rs("title")%>"><img src="<%=rs("img")%>" width="234" height="150" alt="<%= rs("title") %>" onerror="this.src='images/nologo.jpg'"/><p><%= stvalue(rs("title"),21) %></p></a>
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
<div class="title tubiao_C">案例分类 <span>Case Nav</span></div>
<ul class="list_2">
<%=zych_case_class%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
