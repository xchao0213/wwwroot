<!--#include file="../../Include/conn.asp"-->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"-->
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="../css/style.css" media="screen" />
<link rel="stylesheet" type="text/css" href="../css/flexslider.css" media="screen" />
<link rel="stylesheet" type="text/css" href="../css/jquery.jslides.css" media="screen" />
<script type="text/javascript" src="../js/jQuery.js"></script> 
<script type="text/javascript" src="../js/iscroll.js"></script> 
</head>
<body>
<div id="wrapper">
<!--Header Starts-->
<!--#include file="../Include/top.asp"-->  
<section id="main">
<!--Block Ends-->
<div class="block_module paper_bh_white">
<h2>服务项目</h2>
<div id="accordion_menu">
<%
set rs=server.createobject("adodb.recordset") 
exec="select * from project order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">暂无服务项目！</div>"
end if
do while not rs.eof
url="Show.asp?id="&rs("id")&""
img=Replace(rs("img"),"../",dir)
%>
<h5><%=rs("name")%><span class="arrow"></span></h5>
<div class="ac_content">
<p align="center"><img src="<%=img%>" alt="<%=rs("name")%>" width="220" height="120" onerror="this.src='<%=dir%>images/nologo.jpg'"/></p>
<p><%=rs("shuomin")%></p>
<p align="right"><a href="<%=url%>" title="<%=rs("name")%>">MOER</a></p>
</div>
<%rs.movenext
Loop
rs.close
Set rs=nothing%>
</div>
</div>
<!--Block Ends-->
</section>
<!--Section Ends-->
<!--Footer Starts-->
<!--#include file="../Include/bottom.asp" -->
<!--Footer Ends-->
</div>
<!-- JQuery -->
<script src="../js/jquery.min.js"></script>
<!-- JQueryUI -->
<script src="../js/jquery-ui-1.8.16.custom.min.js"></script>
<!--Simple Carousel-->
<script src="../js/jquery.flexslider.js"></script>
<!-- Script Controls -->
<script src="../js/site.js"></script>
<script type="text/javascript">
/* Slideshow Control */		
	jQuery('.flexslider').flexslider();
</script>
</body>
</html>
