<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
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
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>Culture/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<div class="list">
<% 
set culture_fl=server.createobject("adodb.recordset") 
exec="select top 3 * from [culture_fl] order by px_id asc  " 
culture_fl.open exec,conn,1,1 
if culture_fl.eof and culture_fl.bof then
response.Write("&nbsp;�����κη��� !")
end if
do while not culture_fl.eof
if html=0 then
ncurl=""&dir&"Culture/"&Newsfl&""&Separated&""&culture_fl("id")&"."&HTMLName&""
elseif html=1 then
ncurl=""&dir&"Culture/Culturelist.asp?id="&culture_fl("id")&""
end if%>
<div class="dt"><div class="N1">NO</div><div class="N2"><%=culture_fl("title")%></div><div class="N3"><a href="<%= ncurl %>" style="color:#FFF">MORE</a></div></div>
<%set culture=server.createobject("adodb.recordset") 
exec="select top 8 * from [culture] where ssfl="&culture_fl("id")&" order by id desc  " 'top 8����ÿ��Ŀ��ʾ5������
culture.open exec,conn,1,1 
if culture.eof and culture.bof then
response.Write("&nbsp;�˷����������� !")
end if
while not culture.eof
i=(i+1)
data=right("0"&year(culture("data")),4)&"-"&right("0"&month(culture("data")),2)&"-"&right("0"&day(culture("data")),2)
if IsNull(culture("url")) or trim(culture("url")&"")="" then
if html=0 then
url=""&dir&"Culture/"&ShowNews&""&Separated&""&culture("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Culture/ShowCulture.asp?id="&culture("id")&""
end if
else
url=""&culture("url")&""
end if
if culture("img")<>"" then img=culture("img") else img="../images/nologo.jpg"
If newslb<>1 Then %>
<div class="dl"><div class="N1"><%=right("00"&i,3)%></div>
<div class="N2"><a href="<%=url%>"><img src="<%=img%>" width="110" height="70" alt="<%=stvalue(culture("title"),60)%>" onerror="this.src='../images/nologo.jpg'" /></a>
<a href="<%=url%>" class="t" title="<%=stvalue(culture("title"),60)%>"><%=stvalue(culture("title"),60)%></a><p><%=stvalue(Delhtml(culture("body")),100)%></p></div>
<div class="N3"><%=culture("data")%></div></div>
<% Else %>
<div class="dd"><div class="N1"><%=right("00"&i,3)%></div><div class="N2"><a title="<%=culture("title")%>" target="_blank" href="<%=url%>"><%=stvalue(culture("title"),60)%></a></div><div class="N3"><%=culture("data")%></div></div>
<%end if
culture.movenext
wend
culture.close
set culture=nothing
culture_fl.movenext
loop
culture_fl.close
set culture_fl=nothing%>
</div>
<div class="page"></div>
</div>
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_A">��ҵ���� <span>About us</span></div>
<ul class="list_2">
    <li><a href="/" title="��ҳ" class="selected">��ҳ</a></li>
    <li><a href="/About/" title="�߽�����" class="">�߽�����</a></li>
    <li><a href="/News/" title="���ݶ�̬" class="">���ݶ�̬</a></li>
    <li><a href="/about/?id=10" title="������Ŀ">������Ŀ</a></li>
    <li><a href="/Culture/" title="�����Ļ�" class="">�����Ļ�</a></li>
    <li><a href="/Fengfa/" title="�԰��" class="">�԰��</a></li>
    <li><a href="/Contact/" title="��ϵ����" class="">��ϵ����</a></li>    
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
