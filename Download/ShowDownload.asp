<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<%
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
exec="select * from [download] where id="&id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无记录！</div>"
response.End()
end if
set rsc=server.CreateObject("adodb.recordset")
exec="select * from download_fl where id="&rs("ssfl")
rsc.open exec,conn,1,1
if rsc.eof and rsc.bof then
response.Write("&nbsp;暂无记录 !")
end if
fltitle=""&rsc("title")
rsc.close
set rsc=nothing%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("title")%>_<%=zych_home%></title>
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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href=.."/">首页</a><a href="<%=dir%>Download/"><%=zych_class(wurl)%></a><%=rs("title")%></span></div>
<div class="w_650 fl">
<div class="position">
<table width="100%" border="1" cellpadding="6" cellspacing="0" bordercolor="#E1E1E1" style="border-collapse: collapse;line-height:30px;">
<tr>
<td height="30" colspan="2" align="center" bgcolor="#FFFFFF"  style="font-size:16px;font-weight:bold;color:#990000;text-align:center"><%=rs("title")%></td>
</tr>
<tr>
<td width="15%" align="right" bgcolor="#FFFFFF">所属分类：</td>
<td align="left" bgcolor="#FFFFFF"><%=fltitle%></td>
</tr>
<tr>
<td width="15%" align="right" bgcolor="#FFFFFF">程序语言：</td>
<td align="left" bgcolor="#FFFFFF"><%=rs("yuyan")%></td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF">运行平台：</td>
<td align="left" bgcolor="#FFFFFF"><%=rs("yxpt")%></td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF" >推荐等级：</td>
<td align="left" bgcolor="#FFFFFF" ><%=rs("tjdj")%></td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF" >程序大小：</td>
<td align="left" bgcolor="#FFFFFF" ><%=rs("daxiao")%><%=rs("danwei")%></td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF" >下载次数：</td>
<td align="left" bgcolor="#FFFFFF" ><%=rs("js")%> 次</td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF" >程序介绍：</td>
<td align="left" bgcolor="#FFFFFF"><div style="font-size:12px; line-height:20px;"><%=rs("body")%></div></td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF" >发布日期：</td>
<td align="left" bgcolor="#FFFFFF" ><%=rs("data")%></td>
</tr>
<tr>
<td align="right" bgcolor="#FFFFFF" >下载地址：</td>
<td align="left" bgcolor="#FFFFFF"><input name="submit" type="submit" class="input" onClick="window.location.href='<%=dir%>Include/Showbody.asp?action=dowurl&id=<%=rs("id")%>' " value="点击下载" /></td>
</tr>
</table>
</div>
</div>
   
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_D">下载分类 <span>Download Nav</span></div>
<ul class="list_2">
<%=zych_download_fl%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
