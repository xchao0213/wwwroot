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
adminqx="<font color=#FFFFFF>��ͨ����Ա</font>"
else
adminqx="<font color=#FFFFFF>��������Ա</font>"
end if%>
���! [<font color="#FFFFFF"><%=rs("zsname")%></font>]
��ϲ���Ѿ����������ҵ��վ����ϵͳ��̨������û���Ϊ[<%=rs("admin")%>]���ڵļ���Ϊ[<%=adminqx%>]����������Ȩ��!</span></div>
<div id="top">
<div class="logo"></div>
<div class="Navigation"><ul>
<li><a href="../" target="_blank">��վ��ҳ</a></li>
<li><a href="main.asp" target="main">��̨��ҳ</a></li>
<li><a href="admin_xtsz.asp" target="main">��վ����</a></li>
<li><a href="admin_html.asp" target="main">���ɾ�̬</a></li>
<li><a href="add_news.asp" target="main">�������</a></li>
<li><a href="add_case.asp" target="main">��Ӱ���</a></li>
<li><a href="admin_photo.asp"  target="main">������</a></li>
<li><a href="add_download.asp" target="main" >�������</a></li>
<li><a href="Admin_SiteMap.asp" target="main" >SiteMap����</a></li>
<li><a href="LoginOut.asp" target="_parent" >�˳�ϵͳ</a></li>
</ul>
</div>
</div>
</body>
</html>
