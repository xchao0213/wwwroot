<!--#include file="../../Include/conn.asp"-->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/Pager.asp"-->
<!--#include file="../Include/zych_sql.asp"-->
<%'以下为弹出视频
id=request.QueryString("id")
if id="" or not isnumeric(id) then
response.write "<script>alert('警告!请勿尝试非法注入！');window.location.href='index.asp';</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from video where id="&id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此视频！</div>"
response.End()
end if%>
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=rs("title")%>_<%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="../css/style.css" media="screen" />
</head>
<body>
<div id="wrapper">
<!--Header Starts-->
<!--#include file="../Include/top.asp"-->  
<!--Header Ends-->
<!--Section Starts-->
<section id="main">
<!--Block Starts-->
<div class="block_module paper_bh_white">
<div class="page blog">
<a href="javascript:;" id="btn1" class="BttnE" onClick="TopicShow(this,'submenu')"><p>分类</p></a>
<div id="submenu" style="display:none">
<ul><%=zych_video_class%></ul>
<script type="text/javascript">          
function TopicShow(e,TopicID){
e.className=(e.className=="BttnC")?"BttnE":"BttnC"
document.getElementById(TopicID).style.display=(e.className=="BttnC")?"":"none"
}//显示隐藏层
</script></div>
<h1><%=zych_class(murl)%></h1>
<div style="width:100%;margin:0 auto">
<div style="text-indent:5px; font-size:16px; line-height:30px;">视频：<%=rs("title")%></div>
<%if getFileFormat(rs("url"))=".flv" then%>
<script type="text/javascript">
var swf_width=270
var swf_height=200
texts='<%=rs("title")%>'
var files='<%=rs("url")%>'
var swfplay='/images/vcastr2.swf'
var config='0:自动播放|1:连续播放|100:默认音量|0:控制栏位置|2:控制栏显示|0x000033:主体颜色|60:主体透明度|0x66ff00:光晕颜色|0xffffff:图标颜色|0xffffff:文字颜色|<%=Request.ServerVariables("server_name")%>:标题|:logo地址|:结束swf地址'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="'+ swfplay +'"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'">');
document.write('<embed src="'+ swfplay +'" wmode="opaque" FlashVars="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'" menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%else%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="100%" height="auto">
<param name="movie" value="<%=rs("url")%>" />
<param name="quality" value="high" />
<embed src="<%=rs("url")%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="100%" height="auto"></embed>
</object>
<%end if%>
<div style="width:100%;height:30px; background:#000; color:#FFF; list-style:30px">
<div style="float:left;height:30px;line-height:30px; padding-left:5px"><%=rs("title")%></div>
<div style="float:right;text-align:right;line-height:30px; padding-right:5px">点击：<script language="javascript" src="/Include/Click.asp?click=video&id=<%=id%>"></script>次</div>
</div>
<p style="padding:5px; background:#FFF">简介：<%=rs("body")%></p>
</div></div>
</div>
<div class="block_module paper_bh_white">
<div class="list_page"><%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%></div>
</div>
<!--Block Ends-->
</section>

<!--#include file="../Include/bottom.asp" -->
</div>
<!-- JQuery -->
<script src="js/jquery.min.js"></script>
<script src="js/jquery-ui-1.8.16.custom.min.js"></script>
<script src="js/slider.js"></script>
<script src="js/site.js"></script>
</body>
</html>
