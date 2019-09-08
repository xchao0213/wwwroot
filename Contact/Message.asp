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
<meta name="keywords" content="<%=zych_key%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<link rel="stylesheet" type="text/css" href="css/Message.time.css">
<script src="<%=dir%>js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="<%=dir%>js/timeline.message.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>留言记录</b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Contact/">联系我们</a>用户留言</span></div>
<div class="w_650 fl">

<div class="demo">
<%id=request.QueryString("id")
if  not isnumeric(id)  then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from ly where sh=1 order by id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div clacc=""onxx"">暂无留言记录！</div>"
else
rs.PageSize =3
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
end if %>
<div class="history"><div class="more"><a href="/Contact/">去这里说点什么</a></div>
<div class="history-date">
<h2><a class="first" href="#nogo">最新留言</a></h2>
<ul>
<%For i=1 To x%>
<li><h3><%=formatDate(rs("data"),3)%><span><%=formatDate(rs("data"),12)%></span></h3>
<span class="leftcorner"></span>
<dl><h4><span class="fl"><i>F<%=rs("ID")%>#</i><%=rs("name")%></span> <span class="fr">IP：<%=ipvalue(rs("ip"),8)%>&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("data")%></span> </h4>
<p class="con"><%=stvalue(DelHtml(rs("body")),200)%></p>
<h5><b>Reply：</b>
<%
if IsNull(rs("hf")) or trim(rs("hf")&"")="" then
response.Write("<font color=#FF0000>此留言暂未回复！</font>")
else
response.Write(""&rs("hf")&"")
end if%>
</h5>
</dl>
</li>
<%rs.movenext
next
rs.close
set rs=nothing%>
      </ul>
    </div>
  </div>
</div>

</div>
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_C">关于我们 <span>About us</span></div>
<ul class="list_2">
<%=zych_about_list%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
