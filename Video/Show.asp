<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"-->
<%'����Ϊ������Ƶ
id=request.QueryString("id")
if id="" or not isnumeric(id) then
response.write "<script>alert('����!�����ԷǷ�ע�룡');window.location.href='index.asp';</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from video where id="&id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">�޴���Ƶ��</div>"
response.End()
end if%>
<title><%=zych_class(wurl)%>_<%=rs("title")%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link href="<%=dir%>css/style.css" rel="stylesheet" type="text/css" />
<body style="background:#FFF">
<!--������������ -->
<div style="width:700px;margin:0 auto">
<div style="text-indent:5px; font-size:16px; line-height:30px;">��Ƶ��<%=rs("title")%></div>
<%if suffix(rs("url"))=".flv" then%>
<script type="text/javascript">
var swf_width=700
var swf_height=400
texts='<%=rs("title")%>'
var files='<%=rs("url")%>'
var swfplay='/images/vcastr2.swf'
var config='1:�Զ�����|1:��������|100:Ĭ������|0:������λ��|2:��������ʾ|0x000033:������ɫ|60:����͸����|0x66ff00:������ɫ|0xffffff:ͼ����ɫ|0xffffff:������ɫ|<%=Request.ServerVariables("server_name")%>:����|:logo��ַ|:����swf��ַ'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="'+ swfplay +'"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'">');
document.write('<embed src="'+ swfplay +'" wmode="opaque" FlashVars="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'" menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%else
newstr=split(rs("url"),"/id_")'��ַ�ָ�
urlid=split(newstr(1),".")%>
<embed src="http://static.youku.com/qplayer.swf?playMode=mp4&winType=index&VideoIDS=<%=urlid(0)%>&isAutoPlay=false&ShowRelatedVideo=false" type="application/x-shockwave-flash" allowscriptaccess="always" allowfullscreen="true" wmode="opaque" width="700" height="400"></embed>
<%end if%>
<div style="width:700px;height:30px; background:#000; color:#FFF; list-style:30px">
<div style="float:left; width:400px; height:30px;line-height:30px; padding-left:5px"><%=rs("title")%></div>
<div style="float:right; width:100px; text-align:right;line-height:30px; padding-right:5px">�����<script language="javascript" src="/Include/Click.asp?click=video&id=<%=id%>"></script>��</div>
</div>

<p style="padding:5px; background:#FFF">��飺<%=rs("body")%></p>
</div>
</body>