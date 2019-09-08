<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(29)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>在线QQ客服管理</title>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">在线QQ管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<tr>
<td width="6%" align="center" class="td">ID</td>
<td width="15%" height="25" align="center" class="td">客服名称</td>
<td width="23%" align="center" class="td">QQ号</td>
<td width="23%" align="center" class="td">ALT信息</td>
<td width="18%" align="center" class="td">排序</td>
<td width="8%" align="center" class="td">修改</td>
<td width="7%" align="center" class="td">删除</td>
</tr>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from zych_qq order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无qq在线客服!</div>")
else
rs.PageSize =300 '每页记录条数
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
<form action="admin_qq.asp?xiugai=ok" method="post" name="add">
<tr>
<td width="6%" align="center" class="td"><input name="id" type="text" readonly="readonly" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="15%" height="25" align="center" class="td"><input name="bt" type="text" size="20"  value="<%=rs("bt")%>"/></td>
<td width="23%" align="center" class="td"><input name="QQ" type="text" size="20"  value="<%=rs("QQ")%>"/></td>
<td width="23%" align="center" class="td"><input name="alt" type="text" value="<%=rs("alt")%>" size="20"/></td>
<td width="18%" align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td width="8%" align="center" class="td"><input type="submit" name="button2" id="button2" value="修改"  class="btn"/></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_qq.asp?act=del&id=<%=rs("id")%>';}else{history.go(0);}"  class="btn"/></td>
</tr></form>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
</table>    
 </td>
  </tr>
</table>

<div style="margin-top:10px">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="admin_qq.asp?add=ok" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">在线QQ增加</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="16%" height="25" align="right" class="td">客服名称</td>
        <td width="84%"  class="td"><input name="bt" type="text" size="20"  /></td>
      </tr>
      
      <tr>
        <td width="16%" height="25" align="right" class="td">客服QQ号</td>
        <td class="td"><input name="QQ" type="text" size="30"  /></td>
      </tr>
	  <tr>
        <td height="25" align="right" class="td">alt信息</td>
        <td class="td"><input name="alt" type="text" size="30"></td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" size="10"  /> 
          数字越小越靠前。</td>
      </tr>
      
      <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><label>
        <input type="submit" name="button" id="button" value="增加在线客服"  class="btn"/>
        </label></td>
      </tr>
      
    </table></td>
  </tr></form>
</table>

    </div>
</body>
</html>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from zych_qq"
rs.open sql,conn,1,3
bt=request.form("bt")
qq=request.form("qq")
alt=request.form("alt")
px_id=request.form("px_id")
if bt=""  then 
response.Write("<script language=javascript>alert('在线客服名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if qq=""  then 
response.Write("<script language=javascript>alert('QQ号不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert(""错误！排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("bt")=bt
rs("qq")=qq
rs("alt")=alt
rs("px_id")=px_id
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('在线客服增加成功！');window.location.href='admin_qq.asp';</script>" 
end if
%>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id") 
sql="select * from zych_qq where id="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("bt")=request.form("bt") 
rs("qq")=request.form("qq")
rs("alt")=request.form("alt") 
rs("px_id")=request.form("px_id") 
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""当前在线客服修改成功！"");window.location.href='admin_qq.asp';</script>")
end if
%> 
<%
if request("act")="del" then
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('admin_qq.asp');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from zych_qq where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');document.location.href('admin_qq.asp');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('当前在线客服删除成功！');document.location.href('admin_qq.asp');</script>"
end if
end if
%>
