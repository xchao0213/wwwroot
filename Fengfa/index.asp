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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>fengfa/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<div class="list">
<% 
set fengfa_fl=server.createobject("adodb.recordset") 
exec="select top 3 * from [fengfa_fl] order by px_id asc  " 
fengfa_fl.open exec,conn,1,1 
if fengfa_fl.eof and fengfa_fl.bof then
response.Write("&nbsp;�����κη��� !")
end if
do while not fengfa_fl.eof
if html=0 then
ncurl=""&dir&"fengfa/"&Newsfl&""&Separated&""&fengfa_fl("id")&"."&HTMLName&""
elseif html=1 then
ncurl=""&dir&"fengfa/fengfalist.asp?id="&fengfa_fl("id")&""
end if%>
<div class="dt"><div class="N1">NO</div><div class="N2"><%=fengfa_fl("title")%></div><div class="N3"><a href="<%= ncurl %>" style="color:#FFF">MORE</a></div></div>
<%set fengfa=server.createobject("adodb.recordset") 
exec="select top 8 * from [fengfa] where ssfl="&fengfa_fl("id")&" order by id desc  " 'top 8����ÿ��Ŀ��ʾ5������
fengfa.open exec,conn,1,1 
if fengfa.eof and fengfa.bof then
response.Write("&nbsp;�˷����������� !")
end if
while not fengfa.eof
i=(i+1)
data=right("0"&year(fengfa("data")),4)&"-"&right("0"&month(fengfa("data")),2)&"-"&right("0"&day(fengfa("data")),2)
if IsNull(fengfa("url")) or trim(fengfa("url")&"")="" then
if html=0 then
url=""&dir&"fengfa/"&ShowNews&""&Separated&""&fengfa("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Fengfa/ShowFengfa.asp?id="&fengfa("id")&""
end if
else
url=""&fengfa("url")&""
end if
if fengfa("img")<>"" then img=fengfa("img") else img="../images/nologo.jpg"
If newslb<>1 Then %>
<div class="dl"><div class="N1"><%=right("00"&i,3)%></div>
<div class="N2"><a href="<%=url%>"><img src="<%=img%>" width="110" height="70" alt="<%=stvalue(fengfa("title"),60)%>" onerror="this.src='../images/nologo.jpg'" /></a>
<a href="<%=url%>" class="t" title="<%=stvalue(fengfa("title"),60)%>"><%=stvalue(fengfa("title"),60)%></a><p><%=stvalue(Delhtml(fengfa("body")),100)%></p></div>
<div class="N3"><%=fengfa("data")%></div></div>
<% Else %>
<div class="dd"><div class="N1"><%=right("00"&i,3)%></div><div class="N2"><a title="<%=fengfa("title")%>" target="_blank" href="<%=url%>"><%=stvalue(fengfa("title"),60)%></a></div><div class="N3"><%=fengfa("data")%></div></div>
<%end if
fengfa.movenext
wend
fengfa.close
set fengfa=nothing
fengfa_fl.movenext
loop
fengfa_fl.close
set fengfa_fl=nothing%>
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
