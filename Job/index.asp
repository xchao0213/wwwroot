<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
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
<!--[if lt IE 7]>
<script type="text/javascript" src="/images/putaojiayuan.js"></script>
<![endif]-->
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>Job/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from job order by id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">���޼�¼��</div>"
else
rs.PageSize =10
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
end if
For i=1 To x
if html=0 then
url=""&dir&"Job/"&Showjob&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&dir&"Job/Showjob.asp?id="&rs("id")
end if%>
<div class="job_box">
<table width="100%" border="1" cellpadding="6" cellspacing="0" bordercolor="#CECECE" style="border-collapse: collapse">
<tr>
<td width="38%" align="left" bgcolor="#FFFFFF">ְ��<a href="<%=url%>" style="color:#FF0000"><%=stvalue(rs("title"),20)%></a></td>
<td width="30%" align="left" bgcolor="#FFFFFF">�Ա�<%= rs("sex") %></td>
<td width="32%" align="left" bgcolor="#FFFFFF">ѧ����<%= rs("xueli") %></td>
</tr>
<tr>
<td align="left" bgcolor="#FFFFFF">���䣺<%=rs("nn1")%>��<%=rs("nn2")%>��</td>
<td align="left" bgcolor="#FFFFFF">������<%=rs("renshu")%>��</td>
<td align="center" bgcolor="#FFFFFF" style="font-family:'΢���ź�'; font-weight:bold"><a href="<%= url %>" style="font-size:14px">�� ��</a></td>
</tr>
<tr>
<td colspan="3" bgcolor="#FFFFFF" align="left">����:<%=stvalue(DelHtml(rs("body")),300)%></td>
</tr>
</table>
</div>
<%rs.movenext
next%>
<div class="page"><%'������ʾ��ҳ
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%>
</div>
</div>
   
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_J">ְλ�б� <span>Job list</span></div>
<ul class="list_2">
<%= zych_job_list %> 
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
