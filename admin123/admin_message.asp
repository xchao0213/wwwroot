<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(33)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>留言管理</title>
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
<%if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">留言管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="4%" align="center" class="td"><input name="ID" type="checkbox" id="ID"/></td>
<td width="4%" align="center" class="td">ID</td>
<td width="29%" height="25" align="center" class="td">留言标题</td>
<td width="10%" align="center" class="td">留言姓名</td>
<td width="11%" align="center" class="td">是否回复</td>
<td width="12%" align="center" class="td">是否审核</td>
<td width="17%" align="center" class="td">留言时间</td>
<td width="6%" align="center" class="td">回复</td>
<td width="7%" align="center" class="td">删除</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?"> 
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from ly order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无留言!</div>")
else
rs.PageSize =30 '每页记录条数
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
<td width="4%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" />
</td>
<td width="4%" align="center" class="td"><input type="text" readonly="readonly" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="29%" height="25" align="center" class="td"><a href="admin_message.asp?action=admin?id=<%=rs("id")%>" style="color:#003399"><%=left(rs("title"),20)%></a></td>
<td width="10%" align="center" class="td"><%=rs("name")%></td>
<td width="11%" align="center" class="td"><%
if IsNull(rs("hf")) or trim(rs("hf")&"")="" then
response.Write("<font color=#FF0000>[未回复]</font>")
else
response.Write("[已回复]")
end if
%></td>
<td width="12%" align="center" class="td"><%if rs("sh")=1 then
response.Write("[已审核]")
else
response.Write("<font color=#FF0000>[未审核]</font>")
end if%></td>
<td width="17%" align="center" class="td"><%=rs("data")%></td>
<td width="6%" align="center" class="td">
<input type="button" name="Submit3" value="回复" onclick="window.location.href='admin_message.asp?id=<%=rs("id")%>&action=reply' "  class="btn"/></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_message.asp?action=admin&id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">全选</td>
<td colspan="7">
<input type="submit" class="btn" onclick="form.action='?sh=ok';" value="批量审核"/>
<input type="submit" class="btn" onclick="form.action='?sh=no';" value="取消审核"/>
<input type="submit" class="btn" onclick="form.action='?del=ok';" value="批量删除"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table></td>
  </tr>
</table>
<%end if
if request.querystring("action")="reply" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from ly where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<!--------------------以下为添加------------------------------------------------------------------------------------------>
<form  name="add" method="post" action="admin_message.asp?id=<%=rs("id")%>&action=reply&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">查看/回复留言</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8" >
        <td width="16%" height="28" align="right" class="td">留言标题</td>
        <td width="84%"  class="td">&nbsp;<%=rs("title")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td width="16%" height="25" align="right" class="td">留言姓名</td>
        <td class="td">&nbsp;<%=rs("name")%></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td width="16%" height="25" align="right" class="td">QQ</td>
        <td class="td">&nbsp;<%=rs("qq")%></td>
      </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" align="right" class="td">E-Mail</td>
        <td class="td">&nbsp;<%=rs("mail")%></td>
    </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" align="right" class="td">手机</td>
        <td class="td">&nbsp;<%=rs("tel")%></td>
      </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" align="right" class="td">留言时间</td>
        <td class="td">&nbsp;<%=rs("data")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="13" align="right" class="td">留言内容</td>
      <td class="td"><label><%=rs("body")%></label></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="12" align="right" class="td">回复内容</td>
      <td class="td"><textarea name="hf" cols="40" rows="8"><%=rs("hf")%></textarea></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" align="right" class="td">&nbsp;</td>
      <td class="td"><label>
        <input type="submit" name="button" id="button" value="回复留言" class="btn"/>
      </label></td>
    </tr>
    </table></td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from ly where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('删除留言成功！');window.location.href='admin_message.asp?action=admin';</script>"
end if 
%>
<%
if Request.QueryString("sh")="ok" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_message.asp?action=admin';</script>" 
response.end()
end if
sql="update ly set sh=1 where id in ("&Request("id")&")" 
conn.execute(sql)
Response.Write "<script>alert('恭喜!审核成功!');window.location.href='admin_message.asp?action=admin';</script>" 
end if
 
if Request.QueryString("sh")="no" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_message.asp?action=admin';</script>" 
response.end()
end if
sql="update ly set sh=0 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('恭喜!取消审核成功!');window.location.href='admin_message.asp?action=admin';</script>" 
end if


if Request.QueryString("del")="ok" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_message.asp?action=admin';</script>" 
response.end()
end if
sql="delete from ly where id in ("&Request("id")&")"
conn.Execute ( sql )
Response.Write "<script>alert('恭喜!操作成功!');window.location.href='admin_message.asp?action=admin';</script>" 
end if
%>
<% 
if request("xiugai")="ok" then
id=request("id")
hf=request.form("hf")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from ly where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
if hf=""  then 
response.Write("<script language=javascript>alert('请输入回复内容');history.go(-1)</script>") 
response.end 
end if
rs("hf")=hf
rs.update 
rs.close 
response.write "<script>alert('回复成功！');window.location.href='admin_message.asp?action=admin';</script>" 
end if
%> 
