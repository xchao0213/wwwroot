<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/Inc.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(7)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>���ɾ�̬</title>
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
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">���ɾ�̬����</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
<%if html=1 then%>
<tr >
<td height="25" align="center" class="td"><a href="admin_SetSite.asp"><strong>�����ڡ����Ӳ������á��н���̬HTML����Ϊ������</strong></a></td>
</tr>
<%else%>
<tr >
<td height="25" class="td">ע������HTMLҳ�����о�̬ҳ�潫�Զ�������վָ����Ŀ¼�£�-------------------------------------------------------------------\\���ߣ�������ҵ��//-------</td>
</tr>
<tr>
<td height="25" class="td">
<a class="nav" href="html/Admin_html_All.Asp">�������о�̬</a>
<a class="nav" href="html/Admin_html_index.Asp">����Ƶ����ҳ</a>
<a class="nav" href="html/Admin_html_Caes.Asp">����ȫ������</a>
<a class="nav" href="html/Admin_html_Products.Asp">����ȫ����Ʒ</a>
<a class="nav" href="html/Admin_html_news.Asp">����ȫ������</a>
<a class="nav" href="html/Admin_html_Caes.Asp">����ȫ���ĵ�</a>
<a class="nav" href="html/Admin_html_video.Asp">����ȫ����Ƶ</a>
<a class="nav" href="html/Admin_html_team.Asp">����ȫ���Ŷ�</a>
<a class="nav" href="html/Admin_html_Down.Asp">����ȫ������</a>
<a class="nav" href="html/Admin_html_photo.Asp">�������Ƶ��</a>
<a class="nav" href="html/Admin_html_about.Asp">���ɹ�������</a>
<a class="nav" href="html/Admin_html_project.Asp">���ɷ�����Ŀ</a>
</td>
</tr>
</table>
	  
</td>
</tr>
</table>
<%end if%>
</body>
</html>
