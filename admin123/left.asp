<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<% 
function zych_adminleftclass(ID)
set rsl=server.CreateObject("adodb.recordset")
sql="select * from menu where id="&ID&""
rsl.open sql,conn,1,1
Response.Write(rsl("title"))
end function %> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�����б�ҳ��</title>
<script src="js/prototype.lite.js" type="text/javascript"></script>
<script src="js/moo.fx.js" type="text/javascript"></script>
<script src="js/moo.fx.pack.js" type="text/javascript"></script>
<style>
body {font:12px Arial, Helvetica, sans-serif;color: #000;background-color: #EEF2FB;margin: 0px 0px 0px 3px;}
#container {width: 140px;}
H1 {font-size: 12px;margin:0px;cursor: pointer;height: 30px;line-height: 20px;}
H1 a {display:block;color:#FFF;height:30px;text-decoration: none;moz-outline-style: none; background:url(images/bg_list.gif) repeat-x #0099FF;line-height: 30px;text-align: center;margin: 0px; border:1px #015e76 solid;}
.content{height:25px; margin-bottom:2px}
.MM {margin: 0px;padding: 0px;left: 0px;top: 0px;right: 0px;bottom: 0px;clip: rect(0px,0px,0px,0px);}
.MM ul {list-style-type: none;margin: 0px;padding: 0px;display: block;}
.MM li {height:25px;font-family: Arial, Helvetica, sans-serif;line-height: 26px;color:#333333;list-style-type: none;display: block;border:1px #0a82aa solid; border-top:none;margin: 0px;padding: 0px;text-align: center; background:url(images/button_bg.gif) repeat-x}
.MM li a{font-family: Arial, Helvetica, sans-serif;line-height: 25px;color: #333333;overflow: hidden;text-decoration: none;}
.MM li a:hover {line-height: 25px;font-weight: bold;color: #006600;background-image: url(images/menu_bg2.gif);margin: 0px;padding: 0px;height: 25px;width:140px;text-decoration: none;}
</style>
</head>
<body>
<h1 class="type"><a href="javascript:void(0)">��վϵͳ����</a></h1>
<div class="content">
  <ul class="MM">
    <li><a href="admin_xtsz.asp" target="main">ȫ�ֲ�������</a></li>
	<li><a href="admin_SetSite.asp"  target="main">���Ӳ�������</a></li>
    <li><a href="admin_menu.asp?menu=admin" target="main">�����˵�����</a></li>
	<li><a href="admin_count.asp" target="main">��̨��½��¼</a></li>
	<li><a href="admin_administrator.asp?action=admin" target="main">����Ա����</a></li>
	<li><a href="admin_sql.asp?action=config" target="main">SQLע������</a> |<a href="admin_sql.asp" target="main"> ����</a> </li>
	<li><a href="Admin_data.asp"  target="main">�ռ�ռ�����</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">��ҳ����</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_about.asp" target="main">��ӵ�ҳ</a></li>
    <li><a href="admin_about.asp" target="main">����ҳ</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">���Ź���</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_news.asp" target="main">�������</a></li>
    <li><a href="admin_news.asp" target="main">��������</a></li>
	<li><a href="admin_newsfl.asp" target="main">������Ŀ����</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">�����Ļ�</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_culture.asp" target="main">����Ļ�</a></li>
    <li><a href="admin_culture.asp" target="main">�����Ļ�</a></li>
	<li><a href="admin_culture_fl.asp" target="main">�Ļ���Ŀ����</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">�԰��</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_fengfa.asp" target="main">���԰��</a></li>
    <li><a href="admin_fengfa.asp" target="main">����԰��</a></li>
	<li><a href="admin_fengfa_fl.asp" target="main">԰����Ŀ����</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">��������</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_case.asp" target="main">��Ӱ�����Ϣ</a></li>
    <li><a href="admin_case.asp" target="main">��������Ϣ</a></li>
	<li><a href="admin_caseclass.asp" target="main">�����������</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">�Ŷӹ���</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_tuandui.asp?action=add" target="main">���ӳ�Ա</a></li>
    <li><a href="admin_tuandui.asp?action=admin" target="main">�����Ŷ�</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">��Ʒ����</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_products.asp" target="main">��Ӳ�Ʒ</a></li>
    <li><a href="admin_products.asp" target="main">�����Ʒ</a></li>
	<li><a href="admin_products_Bigclass.asp" target="main">��Ʒ�������</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">ר���ؿ�����</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_project.asp?action=add" target="main">���ר���ؿ�</a></li>
    <li><a href="admin_project.asp?action=admin" target="main">����ר���ؿ�</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">�쵼���Ĺ���</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_leader.asp?action=add" target="main">����쵼������Ŀ</a></li>
    <li><a href="admin_leader.asp?action=admin" target="main">�����쵼������Ŀ</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">��Ƶ����</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_video.asp?action=add" target="main">������Ƶ</a></li>
    <li><a href="admin_video.asp?action=admin" target="main">������Ƶ</a></li>
	<li><a href="admin_videoclass.asp" target="main">��Ƶ�������</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">����Ƶ��</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_download.asp" target="main">�������</a></li>
    <li><a href="admin_download.asp" target="main">��������</a></li>
	<li><a href="admin_download_fl.asp" target="main">������Ŀ����</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">������</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_photo.asp" target="main">������</a></li>
    <li><a href="admin_photo.asp" target="main">�������</a></li>
	<li><a href="admin_photoclass.asp" target="main">�����Ŀ����</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">�˲Ź���</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_job.asp?action=add" target="main">������Ƹ</a></li>
    <li><a href="admin_job.asp?action=admin" target="main">������Ƹ</a></li>
	<!--<li><a href="admin_Resume.asp?action=admin" target="main">�鿴����</a></li>-->
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">������</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_ad.asp?action=add&id=1" target="main">���JS���</a></li>
    <li><a href="admin_ad.asp?action=admin" target="main">����JS���</a></li>
	<li><a href="admin_indexflash.asp?action=admin" target="main">������ҳ�õƺ��</a></li>
	<li><a href="admin_flash.asp?action=admin" target="main">����Ƶ���õƺ��</a></li>
  </ul>
</div>


<h1 class="type"><a href="javascript:void(0)">��������</a></h1>
<div class="content">
  <ul class="MM">
    <li><a href="admin_qq.asp" target="main">���߿ͷ�</a></li>
	<li><a href="admin_message.asp?action=admin" target="main">���Թ���</a></li>
	<li><a href="admin_user.asp?action=admin" target="main">��Ա����</a></li>
	<li><a href="admin_orders.asp?action=admin" target="main">��������</a></li>
	<li><a href="admin_backup.asp" target="main">���ݱ���</a></li>
	<li><a href="admin_Attachment.asp" target="main">�ϴ���������</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">��������</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_link.asp?action=add" target="main">�����������</a></li>
    <li><a href="admin_link.asp?action=txt" target="main">������������</a></li>
	<li><a href="admin_link.asp?action=img" target="main">ͼƬ��������</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">��վ�ƹ�</a></h1>
<div class="content">
  <ul class="MM">
  <li><a href="http://www.baidu.com/search/url_submit.html" target="main">�ٶȵ�¼���</a></li>
  <li><a href="http://www.google.com/addurl/" target="main">Google��¼���</a></li>
  <li><a href="http://info.so.360.cn/site_submit.html" target="main">360������¼���</a></li>
  <li><a href="http://www.soso.com/help/usb/urlsubmit.shtml" target="main">SOSO���ѵ�¼���</a></li>
  <li><a href="http://zz.jike.com/submit/genUrlForm" target="main">����������¼���</a></li>
  <li><a href="http://www.sogou.com/feedback/urlfeedback.php" target="main">�ѹ��ύ���</a></li>
  <li><a href="http://tellbot.youdao.com/report" target="main">�е���¼���</a></li>
  <li><a href="http://cn.bing.com/docs/submit.Aspx" target="main">��Ӧ��¼���</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">SEO����</a></h1>
<div class="content">
  <ul class="MM">
   <%seourl=Request.ServerVariables("server_name")%>
	<li><a href="http://seo.chinaz.com/?q=<%=seourl%>/" target="main">SEO�ۺϲ�ѯ</a></li>
	<li><a href="http://mytool.chinaz.com/baidusort.aspx?host=<%=seourl%>" target="main">�ٶ�Ȩ�ز�ѯ</a></li>
    <li><a href="http://seo.addpv.com/BaiDuSort/<%=seourl%>" target="main">�ٶ�������ѯ</a></li>
    <li><a href="http://seo.addpv.com/SiteCheck/<%=seourl%>" target="main">��վSEO��ϱ���</a></li>
	<li><a href="http://seo.addpv.com/CheckShot/<%=seourl%>" target="main">�ٶȿ������</a></li>
	<li><a href="http://seo.addpv.com/SuperLink/<%=seourl%>" target="main">������������</a></li>
	<li><a href="http://seo.addpv.com/LinkDetection/<%=seourl%>" target="main">�������Ӽ�⹤��</a></li>
  </ul>
</div>
<h1><a href="javascript:void(0)">��Ȩ��Ϣ</a></h1>
 <div class="content_1">
  <ul class="MM">
    <li>������ҵ��ҵ������ϵͳ</li>
	<li><a href="tencent://message/?uin=495573637&amp;Site=ziliaok8&amp;Menu=yes" target="_blank">����֧�֣��������ܿƼ�</a></li>
    <li><a href="http://www.shgwzy.com" target="_blank">�������ٷ���վ</a></li>
	<li><a href="tencent://message/?uin=495573637&amp;Site=ziliaok8&amp;Menu=yes" target="_blank"> ����Q Q��1445565423</a></li>
    </ul>
	  </div>
<script type="text/javascript">
		var contents = document.getElementsByClassName('content');
		var toggles = document.getElementsByClassName('type');
	
		var myAccordion = new fx.Accordion(
			toggles, contents, {opacity: true, duration: 400}
		);
		myAccordion.showThisHideOpen(contents[0]);
</script>
</body>
</html>
