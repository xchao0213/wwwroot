<!--�ȵ�����-->
<div class="box"><div class="title tubiao_1">�Ƽ����� <span>News</span></div>
<ul class="list_1">
<%set news=server.createobject("adodb.recordset") 
exec="select top 10 * from  [news]  where tuijian=1 order by id desc  " 
news.open exec,conn,1,1 
if news.eof and news.bof then
response.Write("��������...")
end if
f=1
do while not news.eof
if IsNull(news("url")) or trim(news("url")&"")="" then
dim url
if html=0 then
url=""&Dir&"News/"&ShowNews&Separated&news("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"News/ShowNews.asp?id="&news("id")&""
end if
else
url=""&news("url")&""
end if%> 
<li><a href="<%=url%>"><%=stvalue(news("title"),34)%></a></li>
<%news.movenext 
loop 
news.close
set news=nothing%>
</ul>
<div class="more"><a href="../News">>>�˽�����</a></div>
</div>
<!--��ϵ����-->
<div class="box"><div class="title tubiao_2">��ϵ���� <span>Contact Us</span></div>
<div class="Contact">
<p>ȫ��ͳһ��������</p>
<b><%=zych_tel%></b>
<p>���ǵ�����</p>
<b><%=zych_mail%></b>
</div>
</div>