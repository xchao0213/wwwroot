<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(5)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>后台登陆记录</title>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">后台登陆记录</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" class="td"><input name="ID" type="checkbox" id="ID"/></td>
<td width="4%" align="center" class="td">ID</td>
<td width="25%" height="25" align="center" class="td">登陆用户</td>
<td width="43%" align="center" class="td">登陆IP</td>
<td width="25%" align="center" class="td">登陆时间</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="admin_count.asp?del=checkbox"> 
<%set rs=server.createobject("adodb.recordset") 
exec="select * from [admincount] order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无记录!</div>")
else
rs.PageSize =20 '每页记录条数
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
for i=1 to rs.pagesize%>
<tr>
<td width="3%" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="4%" align="center" class="td"><input  type="text" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="25%" height="25" align="center" class="td"><%=rs("name")%> </td>
<td width="43%" align="center" class="td"><a href="http://www.ip138.com/ips138.asp?ip=<%=rs("ip")%>" target="_blank" title="点击查看IP信息"><%=rs("ip")%></a></td>
<td width="25%" align="center" class="td"><%=rs("dldata")%></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr><td><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
  <td align="center">全选</td>
  <td colspan="3"><input type="submit" name="button" id="button" value="删除选中项"  class="btn"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing%>
</td></tr>
</form>
</table>



    
  </tr>
</table>
</body>
</html>
<%
if Request.QueryString("del")="checkbox" then
if Request("id")="" then
Response.Write "<script>alert('请选择要删除的记录!');window.location.href='admin_count.asp';</script>" 
response.end()
end if
dim sql
sql="delete from admincount where id in ("&Request("id")&")"
conn.Execute ( sql )
conn.close
set conn=nothing
Response.Write "<script>alert('批量删除成功!');window.location.href='admin_count.asp';</script>" 
end if
%>
