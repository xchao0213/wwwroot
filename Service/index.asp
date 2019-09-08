<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
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
<!--[if lt IE 7]>
<script type="text/javascript" src="/images/putaojiayuan.js"></script>
<![endif]-->
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Service/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from project order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无服务项目！</div>"
end if
do while not rs.eof
if html=0 then
url=""&dir&"Service/"&Showproject&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Service/ShowProject.asp?id="&rs("id")&""
end if%>
<div class="Ser_box"><a href="<%=url%>"><img src="<%=rs("img") %>" alt="<%=rs("name")%>" width="140" height="60" /></a>
<div class="Ser_r"><h2><a href="<%=url%>"><%=rs("name")%></a></h2>
<p><%=rs("shuomin")%></p>
</div>
</div>
<%rs.movenext
Loop
rs.close
Set rs=nothing%>
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
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
