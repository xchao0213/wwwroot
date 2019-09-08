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
<h1><%=zych_class(murl)%></h1>
<%
set rs=server.createobject("adodb.recordset") 
exec="select * from [download] order by id desc"
rs.open exec,conn,1,3
if rs.eof then
response.Write "<div style=""padding:10px"">暂无记录！</div>"
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
For i=1 To x %>
<div style="margin:10px auto; line-height:25px">
<table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#DEE7F9" style="text-indent:5px">
<tr>
    <td height="25" colspan="2" bgcolor="#F7F7F7" class="td"><b><%=stvalue(rs("title"),38)%></b></td>
    </tr>
  <tr>
    <td height="25"  bgcolor="#F7F7F7" class="td">大小：<%=rs("daxiao")%><%=rs("danwei")%></td>
<td class="td" width="51%" bgcolor="#F7F7F7"><%=rs("data")%></td>
  </tr>
  <tr>
    <td  class="td"colspan="2" bgcolor="#FFFFFF" style="line-height:23px"><%=stvalue(DelHtml(rs("body")),200)%>...</td>
    
  </tr>
  <tr>
    <td colspan="2"  bgcolor="#F7F7F7" class="td">运行平台：<%=rs("yxpt")%></td>
    </tr>
  <tr>
    <td width="48%"  bgcolor="#F7F7F7" class="td">推荐等级：<%=rs("tjdj")%></td>
    <td align="center" colspan="2" bgcolor="#F7F7F7" class="td"><input type="submit" value="点击下载" class="input" onClick="window.location.href='<%=dir%>Include/Showbody.asp?action=dowurl&id=<%=rs("id")%>' " />
</td>
  </tr>
</table>
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
