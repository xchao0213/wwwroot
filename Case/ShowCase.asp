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
exec="select * from [case] where id="&id
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此案例！</div>"
response.End()
end if
set fl=server.createobject("adodb.recordset") 
exec="select * from case_class where id="&rs("ssfl")&""
fl.open exec,conn,1,1
fl_id=fl("id")
fl_title=fl("title")
if html=0 then
classtitle="<a href="""&dir&"Case/"&Casefl&""&Separated&""&fl("id")&"."&HTMLName&""">"&fl("title")&"</a>"
elseif html=1 then
classtitle="<a href="""&dir&"Case/?id="&fl("id")&""">"&fl("title")&"</a>"
end if
fl.close
set lf=nothing%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rs("title")%>_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<LINK rel="stylesheet" type="text/css" href="../css/sexylightbox.css">
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<SCRIPT src="<%=dir%>js/jquery.easing.1.3.js" type=text/javascript></SCRIPT>
<SCRIPT src="<%=dir%>js/sexylightbox.v2.3.jquery.js" type=text/javascript></SCRIPT>
<SCRIPT language=javascript>
   $(document).ready(function(){
      SexyLightbox.initialize({color:'white', dir: '/images'});
    });
</SCRIPT>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b><%=zych_pd_class(wurl)%></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>Case/"><%=zych_class(wurl)%></a><%=classtitle%></span></div>
<div class="w_650 fl">
<div class="position">
<div class="caes_title"><b><%=rs("title")%></b><p>所属分类：<%=fl_title%></p><p>发布时间：<%=rs("data")%></p>
</div>
<div class="pbox">
<dl>
<%ImagePath=""&rs("ImagePath")&""
if ImagePath="" then
response.Write""
else
ImagePath = split(rs("ImagePath"),"|")
for ii=0 to ubound(ImagePath)
response.Write("<dd><a href='"&ImagePath(ii)&"' rel=""sexylightbox[group1]"" title="""&rs("title")&"""><img src='"&ImagePath(ii)&"' width=""146"" height=""110""/></a></dd>")
next
end if%>
</dl>
</div>
<div class="content-width" style="background:#FFFFFF"><%=rs("body")%></div>
<div id="text" class="text">
<div class="fanye2">
<%
   dim rstmp, nexttitle, prevtitle
   set rstmp=server.CreateObject("adodb.recordset")
   rstmp.open "select top 1 id, title from [case] where id>" & request.QueryString("id") & " order by id asc",conn,1,1
   if not rstmp.eof then
   if html=0 then
   nexttitle= "<a class=""fr"" href="""&dir&"Case/"&Showcase&""&Separated&"" &rstmp(0)&"."&HTMLName&""" >下一组</a>"
   elseif html=1 then
   nexttitle="<a class=""fr"" href="""&dir&"Case/Showcase.asp?id=" & rstmp(0) & """ title="""&stvalue(rstmp(1),60)&""">下一组</a>"
   end if
   else
   nexttitle = "<a class=""fr"">已经没有了！</a>"
   end if
   rstmp.close
   rstmp.open "select top 1 id, title from [case] where id<" &request.QueryString("id")&" order by id desc"
   if not rstmp.eof then
   if html=0 then
   prevtitle= "<a class=""fl"" href="""&dir&"Case/"&Showcase&""&Separated&"" &rstmp(0)&"."&HTMLName&""" >上一组</a>"
   elseif html=1 then
   prevtitle="<a class=""fl"" href="""&dir&"Case/Showcase.asp?id=" & rstmp(0) & """ title="""&stvalue(rstmp(1),60)&""">上一组</a>"
   end if
   else
   prevtitle = "<a class=""fl"">已经没有了！</a>"
   end if
   rstmp.close
   set rstmp=nothing%>
<%=prevtitle%><%=nexttitle%></div>
</div>
</div>
  </div>
    <div class="w_280 fr">
<div class="box">
      <div class="title tubiao_C">案例分类 <span>Case Nav</span></div>
      <ul class="list_2">
        <%=zych_case_class%>
      </ul>
    </div>
      <!--#include file="../Include/left.asp" -->
    </div>
</div>
</div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
