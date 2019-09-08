<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
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
<LINK href="<%=dir%>css/sexylightbox.css" type=text/css rel=stylesheet>
<SCRIPT language=javascript src="<%=dir%>js/jquery-1.4.2.min.js"></SCRIPT>
<SCRIPT src="<%=dir%>js/jquery.easing.1.3.js" type=text/javascript></SCRIPT>
<SCRIPT src="<%=dir%>js/sexylightbox.v2.3.jquery.js" type=text/javascript></SCRIPT>
<SCRIPT language=javascript>
   $(document).ready(function(){
      SexyLightbox.initialize({color:'white', dir: '/images'});
    });
	</SCRIPT>
</head>
<!--[if lt IE 7]>
<script type="text/javascript" src="/images/putaojiayuan.js"></script>
<![endif]-->
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Video/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fr">
<%
id=request.QueryString("id")
if  not isnumeric(id)  then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
if id="" then
exec="select * from video order by id desc"
else
exec="select * from video where ssfl="&id&" order by id desc"
end if
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div class=""onxx"">暂无视频！</div>"
else
rs.PageSize =8
iCount=rs.RecordCount 
iPageSize=rs.PageSize
maxpage=rs.PageCount 
page=request("page")
if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
if page=maxpage then
x=iCount-(maxpage-1)*iPageSize
else
x=iPageSize
end if
end if
For i=1 To x
n="0"&(n+1)
if html<>0 then
url="Show.asp?id="&rs("id")&"&TB_iframe=true&amp;height=450&amp;width=730"
else
url=""&dir&"Video/"&ShowVideo&Separated&rs("id")&"."&HTMLName&"?TB_iframe=true&amp;height=450&amp;width=730"
end if
%>
<div class="video_list">
<div class="video_content"><img src="<%=rs("img")%>" class="video_img"  onerror="this.src='../images/nologo.jpg'" />
<p class="videoTxt"><a href="javascript:;"><%=n%>.<%=rs("title")%> <%=suffix(rs("url")) %></a><em><%=rs("hit")%>次</em></p>
</div>
<div class="video_play"><a href="<%=url%>" rel="sexylightbox[group1]" title="视频：<%=rs("title")%>" target="_blank" class="vLink"></a></div>
</div>
<%rs.movenext
next%>
<div class="page"><%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing
%></div>
  </div>
  <div class="w_280 fl">
    <div class="box">
      <div class="title tubiao_V">视频分类 <span>Video Nav</span></div>
      <ul class="list_2">
        <%=zych_video_class%>
      </ul>
    </div>
    <!--#include file="../Include/left.asp" -->
  </div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
