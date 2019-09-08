<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(14)
id=request.QueryString("id")
key=request.QueryString("key")%>
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
<%
Response.Buffer    =True   
Response.ExpiresAbsolute    =Now()    -    1   
Response.Expires=0   
Response.CacheControl="no-cache"
%>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#cccccc">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">产品管理</div></td>
  </tr>
  <tr><td bgcolor="#FFFFFF"> <div style="float:left; width:120px;padding:10px 10px"> 
<select name="ssfl" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option>---产品大类---</option>
<%set rsk=server.CreateObject("adodb.recordset")
rsk.open "select * from  bigclass order by BigClassID asc",conn,1,1
response.Write("<option value=""?id="">查看全部</option>")
while not rsk.eof
response.Write("<option value=""?id=" & rsk("BigClassID") & """>" & rsk("BigClassName") & "</option>")
rsk.movenext
wend
rsk.close
set rsc=nothing
%></select></div>
<form action="?key=<%=key%>" method="get" name="add"  style="float:left; width:360px;padding:10px 10px">
<b>搜索:</b><input name="key" type="text" value="<%=key%>" size="30" /><input type="submit" class="btn" value="搜索" />
</form>
</td></tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="50" align="center" class="td"><input name="ID" type="checkbox"/></td>
<td width="50" align="center" class="td">产品ID</td>
<td width="75" align="center" class="td">缩略图</td>
<td height="25" class="td">产品名称</td>
<td width="60" align="center" class="td">推荐</td>
<td width="60" align="center" class="td">市场价</td>
<td width="60" align="center" class="td">优惠价</td>
<td width="60" align="center" class="td">已订单数</td>
<td width="60" align="center" class="td">排序</td>
<td width="120" align="center" class="td">发布时间</td>
<td width="60" align="center" class="td">修改</td>
<td width="60" align="center" class="td">删除</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?"> 
<%set rs=server.createobject("adodb.recordset")
if id<> "" then 
where="where BigClassID="&ID&""
elseif key<>"" then
where="where title like '%"&key&"%'"
else 
where=""
end if
exec="select * from products "&where&" order by px_id desc,id desc"  
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
if rs("tuijian")=1 then
tui="<font color=#FF0000>[推荐]</font>"
else
tui="<font color=#000000>[不推]</font>"
end if
if IsNull(rs("img")) or rs("img")="" then
img=""
else
img="<img src="&rs("img")&" width=""70px"" height=""30px""  alt=""缩略图""/>"
end if
if rs("px_id")<>0 then bj="style=""background:#FFC""" else bj=""
%>
<tr <%=bj%>>
<td align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td align="center" class="td"><%=img%></td>
<td class="td"><a href="xiugai_products.asp?id=<%=rs("id")%>"><%=stvalue(DelHtml(rs("title")),50)%></a></td>
<td align="center" class="td"><%=tui%></td>
<td align="center" class="td">￥<%=rs("jiage")%></td>
<td align="center" class="td">￥<%=rs("zkj")%></td>
<td align="center" class="td"><%=rs("dd_id")%></td>
<td align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td align="center" class="td"><%=rs("data")%></td>
<td align="center" class="td"><input type="button" name="Submit3" value="修改" onclick="window.location.href='xiugai_products.asp?id=<%=rs("id")%>' "  class="btn"/></td>
<td align="center" class="td"><input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_products.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext
if rs.eof then exit for
next
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">全选</td>
<td colspan="10"><input type="submit" class="btn" onclick="form.action='?page=<%=page%>&tuijian=ok';" value="推荐"/>
<input type="submit" class="btn" onclick="form.action='?page=<%=page%>&tuijian=no';" value="不推"/>
<input type="submit" class="btn" onclick="form.action='?page=<%=page%>&action=del';" value="删除"/>
<input type="submit" class="btn" onclick="form.action='?page=<%=page%>&action=ordsc';" value="更新排序"/>
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
if Request.QueryString("tuijian")="ok" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="update products set tuijian=1 where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!推荐成功!');window.location.href='admin_products.asp?page="&page&"';</script>" 
end if

if Request.QueryString("tuijian")="no" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="update products set tuijian=0 where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!取消推荐成功!');window.location.href='admin_products.asp?page="&page&"';</script>" 
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from products where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_products.asp?page="&page&"';</script>" 
end if
'=====================================

if Request.QueryString("action")="ordsc" then
depname=request.form("px_id")
myarray=Split(depname,", ") 
set rs=server.createobject("adodb.recordset")
sql="select * from products order by id desc"
rs.open sql,conn,1,3
for i= 0 to Ubound(myarray)
if myarray(i)<>"" then
rs("px_id")=myarray(i)
else
rs("px_id")=0
end if
rs.update
rs.movenext
next
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script>alert('操作成功!');window.location.href='admin_products.asp?page="&page&"';</script>" 
end if
conn.close
set conn=nothing
%>