<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(30)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>订单管理</title>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#cccccc">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">订单管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <form id="form1" name="form1" method="post" action="admin_orders.asp?del=checkbox">
<div style="text-align:left;height:25px; padding:0px 10px 0px 10px;"> 
      <select name="ssfl" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
      <option>---分类查看---</option>
      <option value="?action=admin&lxid=">---全部订单---</option>
      <option value="?action=admin&lxid=1">---新订单---</option>
      <option value="?action=admin&lxid=2">---已处理---</option>
      </select></div>
<thead><tr>
  <td width="48" align="center" class="td"><input name="ID" type="checkbox"/></td>
  <td width="66" align="center" class="td">ID</td>
  <td width="397" align="center" class="td">订单号</td>
  <td width="199" align="center" class="td">订单金额</td>
  <td width="291" height="25" align="center" class="td">订单状态</td>
<td width="344" align="center" class="td">订单时间</td>
<td width="129" align="center" class="td">管理订单</td>
<td width="120" align="center" class="td">删除订单</td>
</tr></thead>
<%id=request.QueryString("lxid") 
set rs=server.createobject("adodb.recordset")
if id<>"" then
exec="select * from [orders] where state="&id&" order by id desc" 
else
exec="select * from [orders] order by id desc"  
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
for i=1 to rs.pagesize %>
<tr>
  <td width="48" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
  <td width="66" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
  <td width="397" align="center" class="td"><a href="?action=xiugai&id=<%=rs("id")%>"><b><%=rs("OrderNo")%></b></a></td>
  <td width="199" align="center" class="td"><%=rs("rmb")%></td>
  <%if rs("state")=1 then
  ddzt="等待买家付款"
  elseif rs("state")=2 then
  ddzt="买家已付款，等待发货"
  elseif rs("state")=3 then
  ddzt="卖家已发货，等待确认收货"
  elseif rs("state")=8 then
  ddzt="交易完成"
  else
  ddzt="订单出现异常"
  end if %> 
  <td width="291" height="25" align="center" class="td"><%=ddzt%></td>
<td width="344" align="center" class="td"><%=rs("data")%></td>
<td width="129" align="center" class="td">
<input type="button" name="Submit3" value="管理订单" onclick="window.location.href='?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="120" align="center" class="td">
<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除该订单？删除后不可恢复!')){window.location.href='admin_orders.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/>
</td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="48" height="30" align="center" style="padding-left:5px;">
<input type="checkbox" name="allbox" onclick="CheckAll()" />
</td>
<td width="66" height="30" align="center" style="padding-left:5px;">全选</td>
<td colspan="2"><label>
  <select name="lxid">
  <option selected="selected" value="">操作类型</option>
  <option value="1">标记为未付款</option>
  <option value="2">标记为已付款</option>
  <option value="3">标记为已发货</option>
  <option value="8">标记为完交易</option>
  <option value="4">订单异常</option>
  </select>
  <input type="submit" name="button" id="button" value="提交"  class="btn"/>
  </label>
  <%heji=conn.execute("Select Sum(rmb) as rmb From orders where state=8")(0)+conn.execute("Select Sum(kuaidi) as kuaidi From orders where state=8")(0)%>
  累计收入:<font style="font-size:24px; color:#F00; font-family:'Trebuchet MS', Arial, Helvetica, sans-serif">￥<%=heji%></font> 元
  
</td>
<td colspan="4"><%call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%></td>
</form>  
    </table></td>
  </tr>
</table>
<% end if %>
<% if request.querystring("action")="xiugai" then
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
end if%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#ccccccc">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">查看订单</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<form name="add" method="post" action="?id=<%=rs("id")%>&xiugai=ok">
	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="stable" >
      <tr >
        <td width="10%" height="28" align="right" class="td">订单编号</td>
        <td width="90%" class="td"><%=rs("OrderNo")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">产品名称</td> 
        <td class="td"><%
set rss=server.createobject("adodb.recordset")
exec="select * from [products] where id in ("&rs("cpid")&")" 
rss.open exec,conn,1,1 
if rss.eof then
response.write ("<a style=""color:#FF0000"">无此产品或此产品被删除!</a>")
else
while not rss.eof
response.Write("<a href=""../Products/ShowProducts.asp?id="&rs("cpid")&""" target=""_blank"">"&rss("title")&"</a><br>")
rss.movenext
wend
rss.close
end if
set rss=nothing
%></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">订购数量</td>
        <td class="td"><%=rs("number")%></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">付款金额</td>
        <td class="td"><input name="rmb" type="text" value="<%=rs("rmb")%>"<% If rs("state")<>1 Then %> style=" background:#E4E4E4" readonly<% End If %> /></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">快递费用</td>
        <td class="td"><input name="kuaidi" type="text" value="<%=rs("kuaidi")%>"<% If rs("state")<>1 Then %> style=" background:#E4E4E4" readonly<% End If %> /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">联系人姓名</td>
        <td class="td"><%=rs("name")%></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">联系电话</td>
        <td class="td"><%=rs("tel")%></td>
      </tr>
   <tr>
      <td height="25" align="right" class="td">收货地址</td>
      <td class="td"><%=rs("address")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">其它说明</td>
      <td class="td"><%=rs("sm")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">提交日期</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">订单状态</td>
      <td class="td"><%
state=rs("state")
if state=1 then
response.Write("<font color=#FF0000>新订单,未处理!</font>")
else
response.Write("<font color=#000000>已经处理!</font>")
end if
%></td>
    </tr>
	<tr>
      <td height="25" align="right" class="td">备注信息</td>
      <td class="td"><textarea name="bz" cols="50" rows="3" ></textarea></td>
    </tr>
	<tr>
      <td height="25" align="right" class="td">物流单号</td>
      <td class="td"><select name="kddm">
            <option selected="selected" value="">快递公司</option>
			<%set rsc=server.CreateObject("adodb.recordset")
            rsc.open "select * from orders_kd order by px_id asc",conn,1,1
            while not rsc.eof
            if rs("kddm")=rsc("kddm") then
            response.Write("<option value="""&rsc("kddm")&""" selected>" & rsc("title") & "</option>")
            else
            response.Write("<option value="""&rsc("kddm")&""">" & rsc("title") & "</option>")
            end if
            rsc.movenext
            wend
            rsc.close
            set rsc=nothing%>
          </select><input name="kddh" type="text" value="<%=rs("kddh")%>" placeholder="请输入快递单号" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">订单状态</td>
      <td class="td"><select name="state">
            <option selected="selected" value="">操作类型</option>
            <option value="1" <%if rs("state")="1" then%>selected<%end if%>>标记为未付款</option>
            <option value="2" <%if rs("state")="2" then%>selected<%end if%>>标记为已付款</option>
            <option value="3" <%if rs("state")="3" then%>selected<%end if%>>标记为已发货</option>
            <option value="8" <%if rs("state")="8" then%>selected<%end if%>>标记为完交易</option>
            <option value="4" <%if rs("state")="4" then%>selected<%end if%>>订单异常</option>
          </select></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">&nbsp;</td>
      <td class="td"><input type="button" name="Submit3" value="返回订单列表" onclick="window.location.href='admin_orders.asp' "  class="btn"/>  <input type="submit" name="button" id="button" value="处理订单"  class="btn"/></td>
    </tr>
    </table>
	</form>
</td>
  </tr>
</table>
<% end if %>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from orders where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('当前订单删除成功！');window.location.href='admin_orders.asp?action=admin';</script>"
end if 
%>
<%
if Request.QueryString("del")="checkbox" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_orders.asp?action=admin';</script>" 
response.end()
end if
dim sql
lx=request.Form("lxid")
if lx="" then
Response.Write "<script>alert('错误!请选择操作类型!');window.location.href='admin_orders.asp?action=admin';</script>" 
response.end
end if
if lx=1 then
sql="update orders set state=1 where id in ("&Request("id")&")" 
conn.execute(sql) 
elseif lx=2 then 
sql="update orders set state=2 where id in ("&Request("id")&")" 
conn.execute(sql)
elseif lx=3 then
sql="update orders set state=3 where id in ("&Request("id")&")" 
conn.execute(sql) 
elseif lx=4 then 
sql="update orders set state=4 where id in ("&Request("id")&")" 
conn.execute(sql) 
else
sql="delete from orders where id in ("&Request("id")&")"
conn.Execute ( sql )
end if
conn.close
set conn=nothing
Response.Write "<script>alert('恭喜!操作成功!');window.location.href='admin_orders.asp?action=admin';</script>" 
end if
%>
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
rs("rmb")=request.form("rmb")
rs("kuaidi")=request.form("kuaidi")
rs("kddm")=request.form("kddm")
rs("kddh")=request.form("kddh")
rs("bz")=bz
rs("state")=state
rs.update 
rs.close 
response.write "<script>alert('订单修改成功！');window.location.href='admin_orders.asp?action=admin';</script>" 
end if
%> 