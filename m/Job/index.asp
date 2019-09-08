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
exec="select * from [job] order by id desc"
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
<table width="100%" border="1" cellpadding="6" cellspacing="0" bordercolor="#CECECE" style="border-collapse:collapse">
<tr>
<td width="38%" align="left" bgcolor="#FFFFFF">职务：<a href="<%=url%>" style="color:#FF0000"><%=stvalue(rs("title"),20)%></a></td>
<td width="30%" align="left" bgcolor="#FFFFFF">性别：<%= rs("sex") %></td>
<td width="32%" align="left" bgcolor="#FFFFFF">学历：<%= rs("xueli") %></td>
</tr>
<tr>
<td align="left" bgcolor="#FFFFFF">年龄：<%=rs("nn1")%>至<%=rs("nn2")%>岁</td>
<td align="left" bgcolor="#FFFFFF">人数：<%=rs("renshu")%>人</td>
<td align="center" bgcolor="#FFFFFF" style="font-family:'微软雅黑'; font-weight:bold">&nbsp;</td>
</tr>
<tr>
<td colspan="3" bgcolor="#FFFFFF" align="left">详情:<%=stvalue(DelHtml(rs("body")),300)%></td>
</tr>
</table>
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
