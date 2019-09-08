<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/page.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"-->
<% id=request.QueryString("id")
if  not isnumeric(id)  then
Response.Write "<script>alert('警告！请勿尝试注入！');history.go(-1);</script>" 
Response.End()
end if
set rsc=server.CreateObject("adodb.recordset")
if id="" then
exec="select * from download_fl"
else
exec="select * from download_fl where id="&id
end if
rsc.open exec,conn,1,1
if rsc.eof and rsc.bof then
response.Write("&nbsp;暂无记录 !")
end if
if id="" then
Downsclass=""
else
Downsclass=""&rsc("title")&"_"
end if
rsc.close
set rsc=nothing
set rs=server.createobject("adodb.recordset") 
if id="" then
exec="select * from download order by id desc"
else
exec="select * from download where ssfl="&id&" order by id desc"
end if
rs.open exec,conn,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=zych_class(wurl)%>_<%= Downsclass %><%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
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
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Download/"><%=zych_class(wurl)%></a></span></div>
<div class="w_650 fl">
<div class="position">
<%if rs.eof then
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
For i=1 To x
%>
<div style="margin:10px auto; line-height:25px">
<table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#DEE7F9" style="text-indent:5px">
  <tr>
    <td class="td" width="41%" height="25" bgcolor="#F7F7F7">
<%if html=0 then
response.write("<a href="""&Dir&"Download/"&ShowDownload&Separated&rs("id")&"."&HTMLName&""" title="&rs("title")&"><b>"&stvalue(rs("title"),38)&"</b></a>")
elseif html=1 then
response.write("<a href="""&Dir&"Download/ShowDownload.asp?id="&rs("id")&""" title="&rs("title")&"><b>"&stvalue(rs("title"),38)&"</a></b>")
end if%>
</td>
<td class="td" width="32%" bgcolor="#F7F7F7">大小：<%=rs("daxiao")%><%=rs("danwei")%></td>
    <td class="td" width="27%" bgcolor="#F7F7F7"><%=rs("data")%></td>
  </tr>
  <tr>
    <td  class="td"colspan="3" bgcolor="#FFFFFF" style="line-height:23px"><%=stvalue(DelHtml(rs("body")),200)%>...</td>
  </tr>
  <tr>
    <td  bgcolor="#F7F7F7" class="td">运行平台：<%=rs("yxpt")%></td>
    <td  bgcolor="#F7F7F7" class="td">推荐等级：<%=rs("tjdj")%></td>
    <td align="center" bgcolor="#F7F7F7" class="td"><input type="submit" value="点击下载" class="input" onClick="window.location.href='<%=dir%>Include/Showbody.asp?action=dowurl&id=<%=rs("id")%>' " />
</td>
  </tr>
</table>
</div>
<%rs.movenext
next%>
</div>

<div class="page"><%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing
%></div>
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
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
