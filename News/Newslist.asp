<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from News where ssfl="& id &" order by id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">�˷������޼�¼��</div>"
response.End()
else
set fl=server.createobject("adodb.recordset") 
exec="select * from news_fl where id="&rs("ssfl")&""
fl.open exec,conn,1,1
flkey=fl("key")
fltitle=fl("title")
if html=0 then
fldh="<a href="""&dir&"News/"&Newsfl&""&Separated&""&fl("id")&"."&HTMLName&""">"&fl("title")&"</a>"
else
fldh="<a href="""&dir&"News/Newslist.asp?id="&fl("id")&""">"&fl("title")&"</a>"
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=zych_class(wurl)%>_<%=zych_home%></title>
<meta name="keywords" content="<%=flkey%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<script type="text/javascript" src="/img/putaojiayuan.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>News/"><%=zych_class(wurl)%></a><%=fldh%></span></div>
<div class="w_650 fl">
<dl class="list">
<div class="dt"><div class="N1">NO</div><div class="N2"><%=fltitle%></div><div class="N3">DATE</div></div>
<%
rs.PageSize =30 
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
For i=1 To x
n=n+1
if IsNull(rs("url")) or trim(rs("url")&"")="" then
if html=0 then
url=""&dir&"News/"&ShowNews&Separated&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"News/Shownews.asp?id="&rs("id")&""
end if
else
url=""&rs("url")&""
end if
if rs("img")<>"" then img=rs("img") else img="../images/nologo.jpg"
If newslb<>1 Then %>
<div class="dl"><div class="N1"><%=right("00"&i,3)%></div>
<div class="N2"><a href="<%=url%>"><img src="<%=img%>" width="110" height="70" alt="<%=stvalue(rs("title"),60)%>" onerror="this.src='../images/nologo.jpg'" /></a>
<a href="<%=url%>" class="t" title="<%=stvalue(rs("title"),60)%>"><%=stvalue(rs("title"),60)%></a><p><%=stvalue(Delhtml(rs("body")),100)%></p></div>
<div class="N3"><%=rs("data")%></div></div>
<% Else %>
<div class="dd"><div class="N1"><%=right("00"&i,3)%></div><div class="N2"><a title="<%=rs("title")%>" target="_blank" href="<%=url%>"><%=stvalue(rs("title"),60)%></a></div><div class="N3"><%=rs("data")%></div></div>
<%end if
rs.movenext
next
end if%>
</dl>
<div class="page"><%'������ʾ��ҳ
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing
%></div>
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
