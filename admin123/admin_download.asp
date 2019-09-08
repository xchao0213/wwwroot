<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(20)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>下载管理</title>
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form.elements.length);i++){ 
  var e=form.elements[i]; 
  if (e.name!="allbox") e.checked=form.allbox.checked; 
 } 
} 
--> 
</script> 
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">下载管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<div style="text-align:left;height:25px; padding:2px 10px 2px 10px;"> 
<select name="ssfl" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option>---分类查看---</option>
<%set rsk=server.CreateObject("adodb.recordset")
rsk.open "select * from download_fl order by id asc",conn,1,1
response.Write("<option value=""?id="">全部下载</option>")
while not rsk.eof
response.Write("<option value=""?id=" & rsk("id") & """>" & rsk("title") & "</option>")
rsk.movenext
wend
rsk.close
set rsc=nothing%>
</select></div>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="50" align="center" class="td"><input type="checkbox"/></td>
<td width="50" align="center" class="td">ID</td>
<td width="70" align="center" class="td">所属分类</td>
<td class="td">名 称</td>
<td width="60" align="center" class="td">大小</td>
<td width="60" align="center" class="td">语言</td>
<td width="60" align="center" class="td">下载次数</td>
<td width="60" align="center" class="td">下载权限</td>
<td width="120" align="center" class="td">时 间</td>
<td width="50" align="center" class="td">修改</td>
<td width="50" align="center" class="td">删除</td>
</tr></thead>
<form id="form" name="form" method="post" action="?del=checkbox">         
<%id= request.QueryString("id") 
set rs=server.createobject("adodb.recordset")
if id<>"" then
exec="select * from download where ssfl="& id &"  order by id desc" 
else
exec="select * from download order by id desc"  
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
for i=1 to rs.pagesize%>
<tr>
<td align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td align="center" class="td"><input type="text" readonly style="text-align:center;width:40px" value="<%=rs("id")%>"/></td>
<td align="center" class="td"><%set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from download_fl",conn,1,1
while not rsc.eof
if rs("ssfl")=rsc("id") then
response.Write("<a href=""?id=" & rsc("id") & """><font color=#003399>[" & rsc("title") & "]</font></a>")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
%></td>
<td class="td"><a href="xiugai_download.asp?id=<%=rs("id")%>"><%=stvalue(DelHtml(rs("title")),50)%></a>  </td>
<td align="center" class="td"><%=rs("daxiao")%><%=rs("danwei")%></td>
<td align="center" class="td"><%=rs("yuyan")%></td>
<td align="center" class="td"><%=rs("js")%></td>
<td align="center" class="td">下载权限</td>
<td align="center" class="td"><%=rs("data")%></td>
<td align="center" class="td"><input type="button" name="Submit3" value="修改" onclick="window.location.href='xiugai_download.asp?id=<%=rs("id")%>'" class="btn"/></td>
<td align="center" class="td"><input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">全选</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?action=del';" value="删除"/>
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
<%if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from download where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('删除成功！');window.location.href='admin_download.asp';</script>"
end if 
%>
<%
if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from download where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('操作成功!');window.location.href='admin_download.asp';</script>" 
end if
%>
