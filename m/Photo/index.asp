<!--#include file="../../Include/conn.asp"-->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/Pager.asp"-->
<!--#include file="../Include/zych_sql.asp"-->
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=zych_class(murl)%>_<%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="../css/style.css" media="screen" />
</head>
<body>
<div id="wrapper">
<!--Header Starts-->
<!--#include file="../Include/top.asp"-->  
<!--Header Ends-->
<!--Section Starts-->
<section id="main">
<!--Block Starts-->
<div class="block_module paper_bh_white">
<div class="page blog">
<a href="javascript:;" id="btn1" class="BttnE" onClick="TopicShow(this,'submenu')"><p>分类</p></a>
<div id="submenu" style="display:none">
<ul><%=zych_Photo_class%></ul>
<script type="text/javascript">          
function TopicShow(e,TopicID){
e.className=(e.className=="BttnC")?"BttnE":"BttnC"
document.getElementById(TopicID).style.display=(e.className=="BttnC")?"":"none"
}//显示隐藏层
</script></div>
<h1><%=zych_class(murl)%></h1>
<%
id=request.QueryString("id")
if  not isnumeric(id)  then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
if id="" then
exec="select * from [Photo] order by id desc"
else
exec="select * from [Photo] where ssfl="&id&" order by id desc"
end if
rs.open exec,conn,1,3
if rs.eof then
response.Write "<div style=""padding:10px"">暂无记录！</div>"
else
rs.PageSize =8
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
img=Replace(rs("img"),"../",dir)%>
<div class="tpic clear">
<a href="Show.asp?id=<%=rs("id")%>"><img src="<%=img%>" width="270" height="150" alt="<%=rs("title")%>" onerror="this.src='<%=dir%>images/nologo.jpg'"/></a>
<p><a href="Show.asp?id=<%=rs("id")%>"><%=rs("title")%></a></p>
<p class="b"><%=stvalue(DelHtml(rs("body")),200) %></p>
</div>
<%rs.movenext
next%>
</div>
</div>
<div class="block_module paper_bh_white">
<div class="list_page"><%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%></div>
</div>
<!--Block Ends-->
</section>

<!--#include file="../Include/bottom.asp" -->
</div>
<!-- JQuery -->
<script src="js/jquery.min.js"></script>
<script src="js/jquery-ui-1.8.16.custom.min.js"></script>
<script src="js/slider.js"></script>
<script src="js/site.js"></script>
</body>
</html>
