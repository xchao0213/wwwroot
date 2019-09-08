<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from tuandui where id="& id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此成员！</div>"
response.End()
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("name")%>_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="../js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../js/superslide.2.1.js"></script>
<script type="text/javascript" src="../js/dropdown.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%= dir %>Team/"><%=zych_class(wurl)%></a></span></div>
<div class="position">
<div class="w_650 fl">
<div class="team_box">
<div class="team_top"><img src="<%=rs("img")%>" width="200" height="136" alt="<%=rs("name")%>" onerror="this.src='../images/nologo.jpg'"/>
<ul class="int">
<li>姓&nbsp;&nbsp;&nbsp;&nbsp;名：<samp><%=rs("name")%></samp></li>
<li>职&nbsp;&nbsp;&nbsp;&nbsp;务：<samp><%=rs("zw")%></samp></li>
<li>专&nbsp;&nbsp;&nbsp;&nbsp;项：<samp><%=rs("unit")%></samp></li>
<li>博&nbsp;&nbsp;&nbsp;&nbsp;客：<samp><%=rs("url")%>【<a href="<%=rs("url")%>" target="_blank">访问</a>】</samp></li>
</ul>
</div>
<div class="team_body"><%=rs("body")%></div>
</div>
</div>
<div class="w_280 fr">
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
