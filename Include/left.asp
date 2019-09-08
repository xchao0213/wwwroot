<!--热点新闻-->
<div class="box"><div class="title tubiao_1">推荐新闻 <span>News</span></div>
<ul class="list_1">
<%set news=server.createobject("adodb.recordset") 
exec="select top 10 * from  [news]  where tuijian=1 order by id desc  " 
news.open exec,conn,1,1 
if news.eof and news.bof then
response.Write("暂无新闻...")
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
<div class="more"><a href="../News">>>了解详情</a></div>
</div>
<!--联系我们-->
<div class="box"><div class="title tubiao_2">联系我们 <span>Contact Us</span></div>
<div class="Contact">
<p>全国统一服务热线</p>
<b><%=zych_tel%></b>
<p>我们的邮箱</p>
<b><%=zych_mail%></b>
</div>
</div>