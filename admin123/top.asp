<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="images/style.css" rel="stylesheet" type="text/css" />
<title>top</title>
<style>
body{font-size:12px;margin:0;padding:0;}
.top {width:100%; height:25px; line-height:25px; background:#0a82aa; text-align:right; color:#FFFFFF; border-bottom:1px #016d89 solid}
#top { width:100%; height:70px;background:url(images/head_bg.gif) repeat-x top; border-bottom:1px #016d89 solid}
#top .logo{width:290px; height:70px; float:left; background:url(images/admin_logo.gif) no-repeat}
#top .Navigation {height:70px; float:left; margin-left:15px;}
#top .Navigation ul {margin:15px 0px 5px 0px; height:25px;} 
#top .Navigation ul li {margin:2px 5px; color: #FFFFFF; border:1px #016d89 solid;line-height:23px;text-align:center; float:left; list-style-type: none; }
#top .Navigation ul li a{padding:0px 8px;color: #FFFFFF;background:url(images/tille.gif) left -23px;display:block;}
#top .Navigation ul li a:hover {background:url(images/tille.gif) left -0px;display:block;}
</style>
</head>
<body>
<div class="top">
<%exec="select * from admin where id="&session("admin")&""
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1
if rs("key")<>0 then
adminqx="<font color=#FFFFFF>普通管理员</font>"
else
adminqx="<font color=#FFFFFF>超级管理员</font>"
end if%>
你好! [<font color="#FFFFFF"><%=rs("zsname")%></font>]
恭喜您已经进入高屋置业网站管理系统后台，你的用户名为[<%=rs("admin")%>]现在的级别为[<%=adminqx%>]请慎用您的权限!</span></div>
<div id="top">
<div class="logo"></div>
<div class="Navigation"><ul>
<li><a href="../" target="_blank">网站首页</a></li>
<li><a href="main.asp" target="main">后台首页</a></li>
<li><a href="admin_xtsz.asp" target="main">网站设置</a></li>
<li><a href="admin_html.asp" target="main">生成静态</a></li>
<li><a href="add_news.asp" target="main">添加新闻</a></li>
<li><a href="add_case.asp" target="main">添加案例</a></li>
<li><a href="admin_photo.asp"  target="main">相册管理</a></li>
<li><a href="add_download.asp" target="main" >添加下载</a></li>
<li><a href="Admin_SiteMap.asp" target="main" >SiteMap生成</a></li>
<li><a href="LoginOut.asp" target="_parent" >退出系统</a></li>
</ul>
</div>
</div>
</body>
</html>
