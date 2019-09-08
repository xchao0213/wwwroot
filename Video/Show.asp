<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
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
<title><%=zych_class(wurl)%>_<%=rs("title")%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link href="<%=dir%>css/style.css" rel="stylesheet" type="text/css" />
<body style="background:#FFF">
<!--顶部导航结束 -->
<div style="width:700px;margin:0 auto">
<div style="text-indent:5px; font-size:16px; line-height:30px;">视频：<%=rs("title")%></div>
<%if suffix(rs("url"))=".flv" then%>
<script type="text/javascript">
var swf_width=700
var swf_height=400
texts='<%=rs("title")%>'
var files='<%=rs("url")%>'
var swfplay='/images/vcastr2.swf'
var config='1:自动播放|1:连续播放|100:默认音量|0:控制栏位置|2:控制栏显示|0x000033:主体颜色|60:主体透明度|0x66ff00:光晕颜色|0xffffff:图标颜色|0xffffff:文字颜色|<%=Request.ServerVariables("server_name")%>:标题|:logo地址|:结束swf地址'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="'+ swfplay +'"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'">');
document.write('<embed src="'+ swfplay +'" wmode="opaque" FlashVars="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'" menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%else
newstr=split(rs("url"),"/id_")'地址分割
urlid=split(newstr(1),".")%>
<embed src="http://static.youku.com/qplayer.swf?playMode=mp4&winType=index&VideoIDS=<%=urlid(0)%>&isAutoPlay=false&ShowRelatedVideo=false" type="application/x-shockwave-flash" allowscriptaccess="always" allowfullscreen="true" wmode="opaque" width="700" height="400"></embed>
<%end if%>
<div style="width:700px;height:30px; background:#000; color:#FFF; list-style:30px">
<div style="float:left; width:400px; height:30px;line-height:30px; padding-left:5px"><%=rs("title")%></div>
<div style="float:right; width:100px; text-align:right;line-height:30px; padding-right:5px">点击：<script language="javascript" src="/Include/Click.asp?click=video&id=<%=id%>"></script>次</div>
</div>

<p style="padding:5px; background:#FFF">简介：<%=rs("body")%></p>
</div>
</body>