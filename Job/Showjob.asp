<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from job where id="& id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此信息！</div>"
response.End()
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("title")%>_<%=zych_home%></title>
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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>job/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<div class="job_box">
<table width="100%" border="1" cellpadding="6" cellspacing="0" bordercolor="#CECECE" style="border-collapse: collapse">
<tr>
<td width="34%" align="left" bgcolor="#FFFFFF">职务: <a href="<%=url%>" style="color:#FF0000"><%=rs("title")%></a></td>
<td width="28%" align="left" bgcolor="#FFFFFF">性别：<%= rs("sex") %></td>
<td width="38%" align="left" bgcolor="#FFFFFF">学历：<%= rs("xueli") %></td>
</tr>
<tr>
<td align="left" bgcolor="#FFFFFF">年龄：<%=rs("nn1")%>至<%=rs("nn2")%>岁</td>
<td align="left" bgcolor="#FFFFFF">人数：<%=rs("renshu")%>人</td>
<td bgcolor="#FFFFFF"><font color="#FF0000">简历投至：<%=zych_mail%></font></td>
</tr>
<tr>
<td colspan="3" align="left">详情:<%=rs("body")%></td>
</tr>
</table>
</div>
</div>
   
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_J">职位列表 <span>Job list</span></div>
<ul class="list_2">
<%= zych_job_list %> 
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
