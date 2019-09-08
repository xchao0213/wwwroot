<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/Inc.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(7)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>生成静态</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<style type="text/css">
<!--
.nav{padding:5px 10px; background:#CCC; margin:15px; display:block; float:left}
.nav:hover{background:#c92323; color:#FFF}
-->
</style>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">生成静态管理</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
<%if html=1 then%>
<tr >
<td height="25" align="center" class="td"><a href="admin_SetSite.asp"><strong>请先在【附加参数配置】中将静态HTML设置为开启！</strong></a></td>
</tr>
<%else%>
<tr >
<td height="25" class="td">注：生成HTML页后所有静态页面将自动存入网站指定的目录下！-------------------------------------------------------------------\\作者：高屋置业人//-------</td>
</tr>
<tr>
<td height="25" class="td">
<a class="nav" href="html/Admin_html_All.Asp">生成所有静态</a>
<a class="nav" href="html/Admin_html_index.Asp">生成频道首页</a>
<a class="nav" href="html/Admin_html_Caes.Asp">生成全部案例</a>
<a class="nav" href="html/Admin_html_Products.Asp">生成全部产品</a>
<a class="nav" href="html/Admin_html_news.Asp">生成全部新闻</a>
<a class="nav" href="html/Admin_html_Caes.Asp">生成全部文档</a>
<a class="nav" href="html/Admin_html_video.Asp">生成全部视频</a>
<a class="nav" href="html/Admin_html_team.Asp">生成全部团队</a>
<a class="nav" href="html/Admin_html_Down.Asp">生成全部下载</a>
<a class="nav" href="html/Admin_html_photo.Asp">生成相册频道</a>
<a class="nav" href="html/Admin_html_about.Asp">生成关于我们</a>
<a class="nav" href="html/Admin_html_project.Asp">生成服务项目</a>
</td>
</tr>
</table>
	  
</td>
</tr>
</table>
<%end if%>
</body>
</html>
