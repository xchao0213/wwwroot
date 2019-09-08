<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Team/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from tuandui where top=1 order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无记录！</div>"
else
while not rs.eof
if html=0 then
url=""&dir&"Team/"&Showteam&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Team/Show.asp?id="&rs("id")&""
end if%>
<div class="team_box">
<div class="team_top"><img src="<%=rs("img")%>" width="200" height="136" alt="<%=rs("name")%>" onerror="this.src='../images/nologo.jpg'"/>
<ul class="int">
<li>姓&nbsp;&nbsp;&nbsp;&nbsp;名：<samp><%=rs("name")%></samp></li>
<li>职&nbsp;&nbsp;&nbsp;&nbsp;务：<samp><%=rs("zw")%></samp></li>
<li>专&nbsp;&nbsp;&nbsp;&nbsp;项：<samp><%=rs("unit")%></samp></li>
<li>博&nbsp;&nbsp;&nbsp;&nbsp;客：<samp><%=rs("url")%>【<a href="<%=rs("url")%>" target="_blank">访问</a>】</samp></li>
</ul>
</div>
<div class="team_int"><%=stvalue(DelHtml(rs("body")),200)%> [<a href="<%= url %>">详情</a>]</div>
</div>
<%rs.movenext
wend	
end if%>
<div style="margin:10px 0px">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from tuandui where top<>1 and pic=1 order by px_id asc,id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无记录！</div>"
else
while not rs.eof
if html=0 then
url=""&dir&"Team/"&Showteam&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Team/Showteam.asp?id="&rs("id")&""
end if%>
<div class="team_a">
<a href="<%=url%>" title="<%=rs("name")%>"><img src="<%=rs("img")%>" width="234" height="150" alt="<%= rs("name") %>" onerror="this.src='images/nologo.jpg'"/><p><%= stvalue(rs("name"),21) %></p></a>
</div>
<%rs.movenext
wend
End If %>
</div>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from tuandui where top<>1 and pic<>1 order by px_id asc,id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write""
else
Response.Write("<ul class=""team_b"">")
while not rs.eof%>
<li> <a href="<%=url%>" title="<%=rs("name")%>"><%= stvalue(rs("name"),21) %></a></li>
<%rs.movenext
wend
end if
Response.Write("</ul>")%>

</div>
  <div class="w_280 fr">
    <div class="box">
      <div class="title tubiao_A">关于我们 <span>About list</span></div>
      <ul class="list_2">
        <%= zych_about_list %>
      </ul>
    </div>
    <!--#include file="../Include/left.asp" -->
  </div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
