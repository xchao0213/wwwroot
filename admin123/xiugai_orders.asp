<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%
if session("qx")=2 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px; "")>您没有管理该模块的权限！</font></div>")
response.End
End If
%>
<% 
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [orders] where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>查看订单</title>
</head>
<body>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">查看订单</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<form name="add" method="post" action="?id=<%=rs("id")%>&xiugai=ok">
	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8" >
        <td height="28" width="16%" class="td">订单编号</td>
        <td width="84%"  class="td"><%=rs("OrderNo")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" width="16%" class="td">产品名称</td>
        <td class="td"><%
set rss=server.createobject("adodb.recordset")
exec="select * from [products] where id="&rs("cpid")&"" 
rss.open exec,conn,1,1 
response.Write("<a href=""../ShowProducts.asp?id="&rs("cpid")&""" target=""_blank"">"&rss("title")&"</a>")
rss.close
set rss=nothing
%></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" width="16%" class="td">订购数量</td>
        <td class="td"><%=rs("number")%></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" class="td">联系人姓名</td>
        <td class="td"><%=rs("name")%></td>
      </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">联系电话</td>
        <td class="td"><%=rs("tel")%></td>
      </tr>
   <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">收货地址</td>
      <td class="td"><%=rs("address")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">其它说明</td>
      <td class="td"><%=rs("sm")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" class="td">提交日期</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">订单状态</td>
      <td class="td"><%
state=rs("state")
if state=1 then
response.Write("<font color=#FF0000>新订单,未处理!</font>")
else
response.Write("<font color=#000000>已经处理!</font>")
end if
%></td>
    </tr>
	<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" class="td">备注信息</td>
      <td class="td">
          <textarea name="bz" cols="50" rows="3" ></textarea></td>
    </tr>
	<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">订单状态</td>
      <td class="td"><select name="state">
            <option selected="selected" value="">操作类型</option>
            <option value="1" <%if rs("state")="1" then%>selected<%end if%>>标记为新订单</option>
            <option value="2" <%if rs("state")="2" then%>selected<%end if%>>标记为已处理</option>
          </select></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" class="td">&nbsp;</td>
      <td class="td"><input type="button" name="Submit3" value="返回订单列表" onclick="window.location.href='admin_orders.asp' "  class="btn"/>  <input type="submit" name="button" id="button" value="处理订单"  class="btn"/></td>
    </tr>
    </table>
	</form>
</td>
  </tr>
</table>
</body>
</html>
<%
if Request.QueryString("xiugai")="ok" then 
id=request("id")
bz=request.form("bz")
state=request.form("state")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from orders where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
rs("bz")=bz
rs("state")=state
rs.update 
rs.close 
response.write "<script>alert('订单修改成功！');window.location.href='admin_orders.asp';</script>" 
end if
%> 