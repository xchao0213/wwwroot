<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from [Products] where id="& id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此记录！</div>"
response.End()
end if
set BigClass=server.createobject("adodb.recordset") 
exec="select * from [bigClass] where BigClassID="&rs("BigClassID")&""
BigClass.open exec,conn,1,1
BigClassName=BigClass("BigClassName")
if html=0 then
BigClassurl=""&Dir&"Products/"&Productsfl&""&Separated&""&rs("BigClassID")&"."&HTMLName&""
else
BigClassurl=""&Dir&"Products/?BigClassID="&rs("BigClassID")&""
end if
if rs("SmallClassID")<>0 then
set Smallclass=server.createobject("adodb.recordset") 
exec="select * from [smallclass] where SmallClassID="&rs("SmallClassID")&""
smallclass.open exec,conn,1,1
SmallClassName=Smallclass("SmallClassName")
	if html=0 then
	SmallClassurl=""&Dir&"Products/"&Productsfl&""&Separated&""&rs("BigClassID")&""&Separated&""&rs("SmallClassID")&"."&HTMLName&""
	else
	SmallClassurl=""&Dir&"Products/?BigClassID="&rs("BigClassID")&"&amp;SmallClassID="&rs("SmallClassID")&""
	end if
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("title")%>_<%=BigClassName%>_<%=zych_home%></title>
<meta name="keywords" content="<%=rs("key")%><%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link type="text/css" href="../css/style.css" rel="stylesheet"/>
<link type="text/css" href="css/css.css" rel="stylesheet">	
<script src="<%=dir%>js/jquery-1.4.2.min.js" type=text/javascript></script>
<SCRIPT src="js/jquery-1.2.6.pack.js" type=text/javascript></SCRIPT>
<SCRIPT src="js/base.js" type=text/javascript></SCRIPT>
<script type="text/javascript">
$(document).ready(function()
{
	$("#secondpane div.p_menu_1").mouseover(function()
    {
	     $(this).css({background:'#dfdfdf',color:'#ff0'}).next("div.p_menu_2").slideDown(300).siblings("div.p_menu_2").slideUp("slow");
         $(this).siblings().css({background:'#ededed',Color:'#FFF'});
	});
});
</script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Products/"><%=zych_class(wurl)%></a><a href="<%=BigClassurl%>"><%=BigClassName%></a><a href="<%= SmallClassurl%>"><%=SmallClassName%></a> </span></div>
<div class="w_650 fr">
<div class="position">
<div id="portop">
<div id="preview">
<div class="jqzoom" id="spec-n1"><img src="<%=rs("img")%>" jqimg="<%=rs("img")%>" width="290"></div>
<div id="spec-n5">
<div class="contro"l id="spec-left"><img src="images/left.gif" /></div>
<div id="spec-list">
<ul class="list-h">
<%ImagePath=rs("ImagePath")
if ImagePath<>"" then
ImagePath=split(ImagePath,"|")
	for ii=0 to ubound(ImagePath)
	Response.Write("<li><a href=""javascript:void(0)""><img src="""&ImagePath(ii)&""" alt="""&rs("title")&"""></a></li>")
	next
else
response.Write "<li><a href=""javascript:void(0)""><img src="""&rs("img")&""" alt="""&rs("title")&"""></a></li>"
end if%>
</ul>
</div>
<div class="control" id="spec-right"><img src="images/right.gif" /></div>
</div>
</div>
<div class="por">
<h1><%=rs("title")%></h1>
<div class="mod"><div class="jiange1">价格</div><div class="jiange2">￥:<%=rs("jiage")%></div><div class="jiange3"><span>￥<%=rs("zkj")%></span></div></div>
<p>所属分类：<%= BigClassName %> > <%= SmallClassName %></p>
<p>产品型号：<%=rs("cpxh")%></p>
<p>产品规格：<%=rs("cpgg")%></p>
<p>发布时间：<%=rs("data")%></p>
<p>推荐星级：<%=rs("xin")%></p>
<div class="gwck">
<div class="a_pay">
<%if pay=0 then %><a href="<%=dir%>alipay/?pay_id=<%=rs("id")%>" class="ali_Buy" target="_blank"><p>在线购买</p></a><%end if%>
<%if rs("taourl")<>"" then %><a href="<%=rs("taourl")%>" class="tao_Buy" target="_blank"><p>在线购买</p></a><%end if%>
</div>
<form  name="form1" method="post" action="ShowProducts.asp?id=<%=rs("ID")%>&action=ok">
<input name="zkj" type="hidden" value="<%=rs("zkj")%>" size="10"/>
<input name="cpsl" type="hidden" class="J_input" id="cpsl" value="1" />
<input name="提交" type="submit" class="paybnt" value="加入购物车" />
</form>
</div>
</div>
</div>
</div>
<div class="spro"><h1>产品详情</h1></div>
<div class="content-width" style="background:#FFFFFF"><%=rs("body")%></div>
<div id="text" class="text">

</div>
</div>

<SCRIPT type=text/javascript>
$(function(){			
 $(".jqzoom").jqueryzoom({
	  xzoom:400,
	  yzoom:380,
	  offset:10,
	  position:"right",
	  preload:1,
	  lens:1
  });
  $("#spec-list").jdMarquee({
	  deriction:"left",
	  width:290,
	  height:56,
	  step:2,
	  speed:4,
	  delay:10,
	  control:true,
	  _front:"#spec-right",
	  _back:"#spec-left"
  });
  $("#spec-list img").bind("mouseover",function(){
	  var src=$(this).attr("src");
	  $("#spec-n1 img").eq(0).attr({
		  src:src.replace("\/n5\/","\/n1\/"),
		  jqimg:src.replace("\/n5\/","\/n0\/")
	  });
	  $(this).css({
		  "border":"2px solid #ff6600",
		  "padding":"1px"
	  });
  }).bind("mouseout",function(){
	  $(this).css({
		  "border":"1px solid #ccc",
		  "padding":"2px"
	  });
  });				
})
</SCRIPT>
<SCRIPT src="js/lib.js" type=text/javascript></SCRIPT>
<SCRIPT src="js/163css.js" type=text/javascript></SCRIPT>

<div class="w_280 fl">
<div class="box">
<div class="title tubiao_P">产品分类 <span>Products Nav</span></div>
<ul class="p_list" id="secondpane">
<%=zych_Products_class%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<%if request.querystring("action")="ok" then
if session("username")<>"" then
  set rst=server.createobject("adodb.recordset")
  sql="select * from [ogwc] where cpid="&id&" and userid="&userid&" order by id desc"
  rst.open sql,conn,1,3
  if not rst.eof then	
	   Call add_echo("本产品你已经被加入购物车","继续选购","javascript:window.history.go(-1)","去购物车并付款","/User/my_addlist.asp")
  else
	  set rs=server.createobject("adodb.recordset")
	  sql="select * from [ogwc]"
	  rs.open sql,conn,1,3
	  cpsl=request.form("cpsl")

	  if cpsl="0" or cpsl=""  then 
	  response.Write("<script language=javascript>alert('产品数量不能为空!');history.go(-1)</script>") 
	  response.end 
	  end if
	  if request.form("zkj")=""  then 
	  response.Write("<script language=javascript>alert('本产品无金额请直接联系我们!');history.go(-1)</script>") 
	  response.end 
	  end if
	  rs.addnew
	  rs("cpid")=id
	  rs("userid")=userid
	  rs("cpsl")=cpsl
	  rs.update
	  rs.close
	  set rs=nothing
	  Call add_echo("您选择的物品已经加入购物车","继续选购","javascript:window.history.go(-1)","去购物车并付款","/User/my_addlist.asp")
  end if
else
	Response.Write "<script>alert('请登陆本站会员');window.location.href='/User/login.asp?tourl="&Request.ServerVariables("HTTP_REFERER")&"';</script>"
end if 
end if%>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
