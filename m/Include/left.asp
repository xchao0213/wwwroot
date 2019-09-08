<!--热点新闻-->
<div class="left_title"><a href="#" title="推荐新闻"><b>推荐新闻</b><span>Recommend</span></a></div>
<div class="left_box">
<dl class="list">
<% 
set news=server.createobject("adodb.recordset") 
exec="select top 10 * from  [news]  where tuijian=1 order by id desc  " 
news.open exec,conn,1,1 
if news.eof and news.bof then
response.Write("&nbsp;暂无新闻 !")
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
<!--联系我们-->
<div class="left_title"><a href="#" title="联系我们"><b>联系我们</b><span>Contact Us</span></a></div>
<div class="left_box">
<ul id="contact">
<li>联系电话1：<%=zych_tel%></li>
<li>联系电话2：<%=zych_fax%></li>
<li>联系邮箱：<%=zych_mail%></li>
<li>联系地址：<%=zych_dz%></li>
<li>官方网站：<%=zych_url%></li>
</ul>
</div>