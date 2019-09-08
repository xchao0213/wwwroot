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
<title>管理列表页面</title>
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
<h1 class="type"><a href="javascript:void(0)">网站系统设置</a></h1>
<div class="content">
  <ul class="MM">
    <li><a href="admin_xtsz.asp" target="main">全局参数设置</a></li>
	<li><a href="admin_SetSite.asp"  target="main">附加参数设置</a></li>
    <li><a href="admin_menu.asp?menu=admin" target="main">导航菜单管理</a></li>
	<li><a href="admin_count.asp" target="main">后台登陆记录</a></li>
	<li><a href="admin_administrator.asp?action=admin" target="main">管理员管理</a></li>
	<li><a href="admin_sql.asp?action=config" target="main">SQL注入设置</a> |<a href="admin_sql.asp" target="main"> 管理</a> </li>
	<li><a href="Admin_data.asp"  target="main">空间占用情况</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">单页管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_about.asp" target="main">添加单页</a></li>
    <li><a href="admin_about.asp" target="main">管理单页</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">新闻管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_news.asp" target="main">添加新闻</a></li>
    <li><a href="admin_news.asp" target="main">管理新闻</a></li>
	<li><a href="admin_newsfl.asp" target="main">新闻栏目管理</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">高屋文化</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_culture.asp" target="main">添加文化</a></li>
    <li><a href="admin_culture.asp" target="main">管理文化</a></li>
	<li><a href="admin_culture_fl.asp" target="main">文化栏目管理</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">奉发园区</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_fengfa.asp" target="main">添加园区</a></li>
    <li><a href="admin_fengfa.asp" target="main">管理园区</a></li>
	<li><a href="admin_fengfa_fl.asp" target="main">园区栏目管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">案例管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_case.asp" target="main">添加案例信息</a></li>
    <li><a href="admin_case.asp" target="main">管理案例信息</a></li>
	<li><a href="admin_caseclass.asp" target="main">案例分类管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">团队管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_tuandui.asp?action=add" target="main">增加成员</a></li>
    <li><a href="admin_tuandui.asp?action=admin" target="main">管理团队</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">产品管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_products.asp" target="main">添加产品</a></li>
    <li><a href="admin_products.asp" target="main">管理产品</a></li>
	<li><a href="admin_products_Bigclass.asp" target="main">产品分类管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">专题特刊管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_project.asp?action=add" target="main">添加专题特刊</a></li>
    <li><a href="admin_project.asp?action=admin" target="main">管理专题特刊</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">领导关心管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_leader.asp?action=add" target="main">添加领导关心项目</a></li>
    <li><a href="admin_leader.asp?action=admin" target="main">管理领导关心项目</a></li>
  </ul>
</div>
<h1 class="type"><a href="javascript:void(0)">视频管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_video.asp?action=add" target="main">增加视频</a></li>
    <li><a href="admin_video.asp?action=admin" target="main">管理视频</a></li>
	<li><a href="admin_videoclass.asp" target="main">视频分类管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">下载频道</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_download.asp" target="main">添加下载</a></li>
    <li><a href="admin_download.asp" target="main">管理下载</a></li>
	<li><a href="admin_download_fl.asp" target="main">下载栏目管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">相册管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="add_photo.asp" target="main">添加相册</a></li>
    <li><a href="admin_photo.asp" target="main">管理相册</a></li>
	<li><a href="admin_photoclass.asp" target="main">相册栏目管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">人才管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_job.asp?action=add" target="main">发布招聘</a></li>
    <li><a href="admin_job.asp?action=admin" target="main">管理招聘</a></li>
	<!--<li><a href="admin_Resume.asp?action=admin" target="main">查看简历</a></li>-->
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">广告管理</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_ad.asp?action=add&id=1" target="main">添加JS广告</a></li>
    <li><a href="admin_ad.asp?action=admin" target="main">管理JS广告</a></li>
	<li><a href="admin_indexflash.asp?action=admin" target="main">管理首页幻灯横幅</a></li>
	<li><a href="admin_flash.asp?action=admin" target="main">管理频道幻灯横幅</a></li>
  </ul>
</div>


<h1 class="type"><a href="javascript:void(0)">其它管理</a></h1>
<div class="content">
  <ul class="MM">
    <li><a href="admin_qq.asp" target="main">在线客服</a></li>
	<li><a href="admin_message.asp?action=admin" target="main">留言管理</a></li>
	<li><a href="admin_user.asp?action=admin" target="main">会员管理</a></li>
	<li><a href="admin_orders.asp?action=admin" target="main">订单管理</a></li>
	<li><a href="admin_backup.asp" target="main">数据备份</a></li>
	<li><a href="admin_Attachment.asp" target="main">上传附件管理</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">友情链接</a></h1>
<div class="content">
  <ul class="MM">
	<li><a href="admin_link.asp?action=add" target="main">添加友情链接</a></li>
    <li><a href="admin_link.asp?action=txt" target="main">文字友情链接</a></li>
	<li><a href="admin_link.asp?action=img" target="main">图片友情链接</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">网站推广</a></h1>
<div class="content">
  <ul class="MM">
  <li><a href="http://www.baidu.com/search/url_submit.html" target="main">百度登录入口</a></li>
  <li><a href="http://www.google.com/addurl/" target="main">Google登录入口</a></li>
  <li><a href="http://info.so.360.cn/site_submit.html" target="main">360搜索登录入口</a></li>
  <li><a href="http://www.soso.com/help/usb/urlsubmit.shtml" target="main">SOSO搜搜登录入口</a></li>
  <li><a href="http://zz.jike.com/submit/genUrlForm" target="main">即刻搜索登录入口</a></li>
  <li><a href="http://www.sogou.com/feedback/urlfeedback.php" target="main">搜狗提交入口</a></li>
  <li><a href="http://tellbot.youdao.com/report" target="main">有道登录入口</a></li>
  <li><a href="http://cn.bing.com/docs/submit.Aspx" target="main">必应登录入口</a></li>
  </ul>
</div>

<h1 class="type"><a href="javascript:void(0)">SEO管理</a></h1>
<div class="content">
  <ul class="MM">
   <%seourl=Request.ServerVariables("server_name")%>
	<li><a href="http://seo.chinaz.com/?q=<%=seourl%>/" target="main">SEO综合查询</a></li>
	<li><a href="http://mytool.chinaz.com/baidusort.aspx?host=<%=seourl%>" target="main">百度权重查询</a></li>
    <li><a href="http://seo.addpv.com/BaiDuSort/<%=seourl%>" target="main">百度排名查询</a></li>
    <li><a href="http://seo.addpv.com/SiteCheck/<%=seourl%>" target="main">网站SEO诊断报告</a></li>
	<li><a href="http://seo.addpv.com/CheckShot/<%=seourl%>" target="main">百度快照诊断</a></li>
	<li><a href="http://seo.addpv.com/SuperLink/<%=seourl%>" target="main">超级外链工具</a></li>
	<li><a href="http://seo.addpv.com/LinkDetection/<%=seourl%>" target="main">友情链接检测工具</a></li>
  </ul>
</div>
<h1><a href="javascript:void(0)">版权信息</a></h1>
 <div class="content_1">
  <ul class="MM">
    <li>高屋置业企业网管理系统</li>
	<li><a href="tencent://message/?uin=495573637&amp;Site=ziliaok8&amp;Menu=yes" target="_blank">技术支持：逸丽智能科技</a></li>
    <li><a href="http://www.shgwzy.com" target="_blank">点击进入官方网站</a></li>
	<li><a href="tencent://message/?uin=495573637&amp;Site=ziliaok8&amp;Menu=yes" target="_blank"> 技术Q Q：1445565423</a></li>
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
