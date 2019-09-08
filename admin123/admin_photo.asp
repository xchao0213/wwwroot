<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(22)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>相册管理</title>
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
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">相册管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<div style="text-align:left;height:25px; padding:2px 10px 2px 10px;"> 
<select name="ssfl" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option>---分类查看---</option>
<%set rsk=server.CreateObject("adodb.recordset")
rsk.open "select * from photo_class order by id asc",conn,1,1
response.Write("<option value=""?id="">全部相册</option>")
while not rsk.eof
response.Write("<option value=""?id=" & rsk("id") & """>" & rsk("title") & "</option>")
rsk.movenext
wend
rsk.close
set rsc=nothing
%></select></div>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="50" align="center" class="td"><input type="checkbox" /></td>
<td width="50" align="center" class="td">ID</td>
<td width="70" align="center" class="td">分类</td>
<td width="70" align="center" class="td">缩略图</td>
<td height="25" class="td">名称</td>
<td width="60" align="center" class="td">推荐</td>
<td width="120" align="center" class="td">时 间</td>
<td width="60" align="center" class="td">修改</td>
<td width="60" align="center" class="td">删除</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="admin_photo.asp?del=checkbox"> 
<%id= request.QueryString("id") 
set rs=server.createobject("adodb.recordset")
if id<>"" then
exec="select * from photo where ssfl="& id &"  order by id desc" 
else
exec="select * from photo order by id desc"  
end if
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
for i=1 to rs.pagesize
if IsNull(rs("img")) or rs("img")="" then img="" else img="<img src="&rs("img")&" width=""70px"" height=""30px""  alt=""缩略图""/>"
if rs("tuijian")=1 then tui="<font color=#FF0000>[推荐]</font>" else tui="[不推]"
%>
<tr>
<td align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td align="center" class="td"><input type="text" readonly style="text-align:center;width:40px" value="<%=rs("id")%>"/></td>
<td align="center" class="td">
<%set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from photo_class",conn,1,1
while not rsc.eof
if rs("ssfl")=rsc("id") then
response.Write("<a href=""?id=" & rsc("id")&"""><font color=#003399>[" & rsc("title") & "]</font></a>")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing%></td>
<td align="center" class="td"><%=img%></td>
<td class="td"><a href="xiugai_photo.asp?id=<%=rs("id")%>"><%=stvalue(DelHtml(rs("title")),40)%></a></td>
<td align="center" class="td"><%=tui%></td>
<td align="center" class="td"><%=rs("data")%></td>
<td align="center" class="td"><input type="button" name="Submit3" value="修改" onclick="window.location.href='xiugai_photo.asp?id=<%=rs("id")%>'"  class="btn"/></td>
<td align="center" class="td">
<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_photo.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<% rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">全选</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?tuijian=ok';" value="推荐"/>
  <input type="submit" class="btn" onclick="form.action='?tuijian=no';" value="不推"/>
  <input type="submit" class="btn" onclick="form.action='?action=del';" value="删除"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>
</td>
  </tr>
</table>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from photo where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('当前照片删除成功！');window.location.href='admin_photo.asp';</script>"
end if 


if Request.QueryString("tuijian")="ok" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="update [photo] set tuijian=1 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('恭喜!推荐成功!');window.location.href='admin_photo.asp';</script>" 
end if

if Request.QueryString("tuijian")="no" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="update [photo] set tuijian=0 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('恭喜!取消推荐成功!');window.location.href='admin_photo.asp';</script>" 
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [photo] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_photo.asp';</script>" 
end if
%>
