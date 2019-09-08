<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%
if session("qx")=2 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px; "")>您没有管理该模块的权限！</font></div>")
response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>产品管理</title>
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script>

</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #666">
  <tr>
    <td valign="top">
<table width="100%" height="30" border="0" cellpadding="0" cellspacing="0" background="images/bg_list.gif">
  <tr>
    <td><div style="margin:8px; font-weight:bold; font-size:14px; color:#fff">搜索结果</div></td>
  </tr>
</table>

<div class="right_body">
<%	
t=request.QueryString("t") 
key=request.QueryString("key") 
if t="" then
Response.Write("<script>alert('请选择要搜索的栏目!');history.back();</script>")
Response.End()
end if
if key="" then
Response.Write("<script>alert('请输入关键词!');history.back();</script>")
Response.End()
end if
set rs=server.createobject("adodb.recordset") 
if t=1 then
exec="select * from [news] where title like '%"&key&"%'  order by id desc  " 
elseif t=2 then
exec="select * from [Products] where title like '%"&key&"%'  order by id desc  " 
else
exec="select * from [download] where title like '%"&key&"%'  order by id desc  " 
end if
if t=4 then
exec="select * from [case] where title like '%"&key&"%'  order by id desc  " 
end if
rs.open exec,conn,1,1 
if rs.eof then
response.Write "&nbsp;没有搜索到相关内容！"
else
rs.PageSize =10 '每页记录条数
iCount=rs.RecordCount '记录总数
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
For i=1 To x%>
<%if t=1 then%>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td width="55%" class="td">· <a href="ShowNews.asp?id=<%=rs("id")%>" style=" font-size:14px" target="_blank"><%=stvalue(rs("title"),50)%></a> 浏览：<%=rs("hit")%></td>
          <td width="45%" class="td"><div align="left"><font color="#009900"><%=rs("data")%></font></div></td>
        </tr>
      </table>
<%end if%>
<%if t=2 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
  
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
        <form id="form1" name="form1" method="post" action="soproducts.asp?del=checkbox"> 
      <tr>
 <td width="197%" colspan="3">       </td>
      </tr>
      <tr onMouseOver="style.backgroundColor='#EEEEEE'" onMouseOut="style.backgroundColor='<%=bg%>'" bgcolor="<%=bg%>">
        <td colspan="4" ><table width="100%" border="0" cellpadding="5" cellspacing="0">
            <tr>
              <td width="6%" class="td"><b>ID：</b><%=rs("id")%></td>
              <td width="28%" height="25" class="td">
<b>名称：</b>
<a href="xiugai_products.asp?id=<%=rs("id")%>"><%=rs("title")%></a>  
<%
if rs("tuijian")=1 then
response.Write("<font color=#FF0000>[推荐]</font>")
else
end if
%></td>
              <td width="8%" class="td"><b>点击：</b><%=rs("hit")%></td>
              <td width="26%" class="td"><b>分类：</b>
<% 
SmallClassID=request.QueryString("SmallClassID")
if not isnumeric(SmallClassID) then
Response.Write "<script>alert('错误的参数！');history.go(-1);</script>" 
Response.End()
end if
%>
<%
exec="select * from [smallclass] where SmallClassID="&rs("SmallClassID")&""
set rss=server.createobject("adodb.recordset") 
rss.open exec,conn,1,1
response.write (""&rss("SmallClassName")&"")
%>
			  </td>
              <td width="16%" class="td"><b>时间：</b><%=rs("data")%></td>
              <td width="7%" class="td"><div align="center">
                <input type="button" name="Submit3" value="修改" onClick="window.location.href='xiugai_products.asp?id=<%=rs("id")%>' "  class="btn"/>
              </div></td>
              <td width="9%" class="td">
                  
                  <div align="center">
                    <input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_products.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/>
                  </div></td>
            </tr>
          </table>            </td>
      </tr>
      </form>
		    
    </table></td>
  </tr>
</table>
<%end if%>
<%if t=3 then%>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td width="55%" class="td">· <a href="ShowDownload.asp?id=<%=rs("id")%>" title="<%=rs("title")%>"><%=stvalue(rs("title"),20)%></a></td>
          <td width="45%" class="td"><div align="left"><font color="#009900"><%=rs("data")%></font></div></td>
        </tr>
      </table>
<%end if%>
<%if t=4 then%>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
        <tr>
          <td width="55%" class="td">· <a href="ShowAnli.asp?id=<%=rs("id")%>" title="<%=rs("title")%>"><%=stvalue(rs("title"),20)%></a></td>
          <td width="45%" class="td"><div align="left"><font color="#009900"><%=rs("data")%></font></div></td>
        </tr>
      </table>
<%end if%>

<%rs.movenext
next
end if
%>
</div>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<form action="soproducts.asp" method="get" name="add">
<b>搜索:</b>
      <input name="key" type="text" value="<%=key%>" size="30" />
      <input name="t" type="hidden" id="t" value="2" /><input type="submit" name="button" id="button" value="搜索" />
</form>
	</td>
    <td>
<%'以下显示分页
 call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%>
	</td>
  </tr>
</table>
</td>
  </tr>
</table>
</body>
</html>