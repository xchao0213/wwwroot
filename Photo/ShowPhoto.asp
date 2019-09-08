<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from [Photo] where id="& id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此图片</div>"
response.End()
end if
set daohang=server.createobject("adodb.recordset") 
exec="select * from Photo_class where id="&rs("ssfl")&""
daohang.open exec,conn,1,1
fltitle=daohang("title")%>
<%
   dim rstmp, nexttitle, prevtitle, nexttitle_a, prevtitle_a
   set rstmp=server.CreateObject("adodb.recordset")
   rstmp.open "select top 1 id, title from [photo] where ssfl="&rs("ssfl")&" and id>" & id & " order by id asc",conn,1,1
   if not rstmp.eof then
   	if html<>0 then
	url=""&dir&"Photo/ShowPhoto.asp?id="&rstmp(0)&""
	else
	url=""&dir&"Photo/"&showPhoto&Separated&rstmp(0)&"."&HTMLName&""
	end if
   nexttitle="<a class=""pr"" href="""&url&""">下一组</a>"
   else
   nexttitle = "<a class=""pr"" title=""本类别已经没有了"">下一组</a>"
   end if
   
   if not rstmp.eof then
    if html<>0 then
	url=""&dir&"Photo/ShowPhoto.asp?id="&rstmp(0)&""
	else
	url=""&dir&"Photo/"&showPhoto&Separated&rstmp(0)&"."&HTMLName&""
	end if
   nexttitle_a="<a class=""next""  href="""&url&"""><small>下一组</small></a>"
   else
   nexttitle_a = "<a class=""next"" title=""本类别已经没有了""><small>下一组</small></a>"
   end if
   rstmp.close
   rstmp.open "select top 1 id, title from [photo] where ssfl="&rs("ssfl")&" and id<" & id & " order by id desc"
   if not rstmp.eof then
   prevtitle="<a class=""pl"" href=""Show.asp?" & rstmp(0) & ".html"">上一组</a>"
   else
   prevtitle = "<a class=""pl"" title=""本类别已经没有了"">上一组</a>"
   end if
   if not rstmp.eof then
   prevtitle_a="<a class=""prev"" href=""Show.asp?" & rstmp(0) & ".html""><small>上一组</small></a>"
   else
   prevtitle_a = "<a class=""prev"" title=""本类别已经没有了""><small>上一组</small></a>"
   end if
   rstmp.close
   set rstmp=nothing
   listtitle="<a href=""../photo/?"&rs("ssfl")&"-1.html"" class=""pc"">返回列表</a>"
   %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("title")%>_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<LINK rel="stylesheet" type="text/css" href="../css/sexylightbox.css">
<link rel="stylesheet" type="text/css" href="assets/imageflow.packed.css" />
<script type="text/javascript" src="<%=dir%>js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<script type="text/javascript" src="assets/index_bnaner_case.js"></script>
<SCRIPT src="<%=dir%>js/jquery.easing.1.3.js" type=text/javascript></SCRIPT>
<SCRIPT src="<%=dir%>js/sexylightbox.v2.3.jquery.js" type=text/javascript></SCRIPT>
<SCRIPT language=javascript>
   $(document).ready(function(){
      SexyLightbox.initialize({color:'white', dir: '/images'});
    });
</SCRIPT>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Photo/"><%=zych_class(wurl)%></a><%=fltitle%></span></div>
<div class="w_650 fl">
<div id="Case_nr">
<div class="int">
<img src="<%=rs("img")%>" alt="<%= rs("title") %>" width="172" height="200" />
<div class="details"><h1><%=stvalue(rs("title"),26)%></h1><p>所属分类：<%=fltitle%></p><p>上传时间：<%= rs("data") %></p>
<p>浏览人气：已有 <font color="#d84b59"><script src="<%=dir%>Include/Click.asp?click=News&id=<%=id%>"></script></font> 人浏览</p>
<p><!-- Baidu Button BEGIN -->
<div id="bdshare" class="bdshare_t bds_tools get-codes-bdshare">
<a class="bds_qzone"></a>
<a class="bds_tsina"></a>
<a class="bds_tqq"></a>
<a class="bds_renren"></a>
<a class="bds_t163"></a>
<span class="bds_more">更多</span>
<a class="shareCount"></a>
</div>
<script type="text/javascript" id="bdshare_js" data="type=tools&amp;uid=542701" ></script>
<script type="text/javascript" id="bdshell_js"></script>
<script type="text/javascript">
document.getElementById("bdshell_js").src = "http://bdimg.share.baidu.com/static/js/shell_v2.js?cdnversion=" + Math.ceil(new Date()/3600000)
</script>
<!-- Baidu Button END --></p>
</div>
<div class="pnbox"><%= prevtitle_a %> <%= nexttitle_a %> </div>
</div>
</div>
<div class="ztzs_01">
<div class="show_support" id="product_ding_1427" onClick="ding('product',1427)" style="cursor:pointer;">
<div style="clear:both;"></div>
</div>
<div style="float:left; padding-left:0px;" class="ztzs_01">
Theme <span style="font-family:microsoft yahei">图组展示</span>  <span style="font-size:12px; font-weight:normal;color:rgb(255,215,0);">【点击下图全屏展示】</span>
</div>
</div>
<div class="banner">
<div class="mod-banner">
<div style="visibility: visible; overflow:hidden; position:relative; z-index: 2; left: 0px; width: 660px;" id="focus_img">
<div class="pcont" id="ISL_Cont_1">
<div class="ScrCont">
<div id="List1_1">
<%ImagePath = split(rs("ImagePath"),"|")
for ii=0 to ubound(ImagePath)
n=n+1
Response.Write"<dl><a style=""width:144px;height:115px;overflow:hidden;display:block;"" href="""&ImagePath(ii)&""" rel=""sexylightbox[group1]"" title="""&rs("title")&"""><img src='"&ImagePath(ii)&"' width=""144"" height=""115""/></a></dl>"&vbcrlf
next%></div>
<div id="List2_1"></div>
</div></div> 
</div>
<a onMouseUp="ISL_StopDown_1()" id="btn_focus_prev" onMouseDown="ISL_GoDown_1()" onMouseOut="ISL_StopDown_1()" href="javascript:void(0);" target="_self"></a>
<a onMouseUp="ISL_StopUp_1()" id="btn_focus_next" onMouseDown="ISL_GoUp_1()" onMouseOut="ISL_StopUp_1()" href="javascript:void(0);" target="_self"></a>
</div>
</div>
<script type="text/javascript">
picrun_ini();//-->JS调用
$(function(){
  $("#focus_img dl ").each(function(){
    $(this).hover(function(){
        $(this).find('img').animate({"opacity":"0.5"},'fast');
    },function(){
        $(this).find('img').animate({"opacity":"1"},'fast');
    })
  })
 });
</script>
<div class="content-width" style="background:#FFFFFF"><%=rs("body")%></div>
</div>

<div class="w_280 fr">
<div class="box">
<div class="title tubiao_P">相册分类 <span>photo Nav</span></div>
<ul class="list_2">
<%=zych_photo_class%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
  
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
