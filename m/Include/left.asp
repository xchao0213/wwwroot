<!--�ȵ�����-->
<div class="left_title"><a href="#" title="�Ƽ�����"><b>�Ƽ�����</b><span>Recommend</span></a></div>
<div class="left_box">
<dl class="list">
<% 
set news=server.createobject("adodb.recordset") 
exec="select top 10 * from  [news]  where tuijian=1 order by id desc  " 
news.open exec,conn,1,1 
if news.eof and news.bof then
response.Write("&nbsp;�������� !")
end if
f=1
do while not news.eof
if IsNull(news("url")) or trim(news("url")&"")="" then
dim url
if html=0 then
url=""&Dir&"News/"&ShowNews&""&Separated&""&news("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"News/Shownews.asp?id="&news("id")&""
end if
else
url=""&news("url")&""
end if
%> 
<dt><a href="<%=url%>"><%=stvalue(news("title"),30)%></a></dt>
<dd>DATA:<%=stvalue(news("data"),60)%></dd>
<%
if f mod 2 =0 then
response.Write ""
end if 
f=f+1 
news.movenext 
loop 
news.close
set news=nothing
%>
</dl>
</div>
<!--��ϵ����-->
<div class="left_title"><a href="#" title="��ϵ����"><b>��ϵ����</b><span>Contact Us</span></a></div>
<div class="left_box">
<ul id="contact">
<li>��ϵ�绰1��<%=zych_tel%></li>
<li>��ϵ�绰2��<%=zych_fax%></li>
<li>��ϵ���䣺<%=zych_mail%></li>
<li>��ϵ��ַ��<%=zych_dz%></li>
<li>�ٷ���վ��<%=zych_url%></li>
</ul>
</div>