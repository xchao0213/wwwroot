<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if id="" or not isnumeric(id) then
response.write "<script>alert('����!�����ԷǷ�ע�룡');window.location.href='index.asp';</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from leader where id="&id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">�޴˷�����Ŀ��</div>"
response.End()
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("name")%>_<%=zych_home%></title>
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
<div class="bar"><b><%=zych_class(wurl)%></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>Service/"><%=zych_class(wurl)%></a><%=rs("name")%></span></div>
<div class="w_650 fl">
<div class="Ser_show"><h2><%=rs("name")%></h2>
<%=rs("body")%>
</div>
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

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
