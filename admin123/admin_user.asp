<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<!--#include file="md5.Asp" -->
<%call chkAdmin(28)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>会员管理</title>
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
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">会员管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" /></td>
<td width="5%" align="center" class="td">ID</td>
<td width="17%" height="25" align="center" class="td">会员帐号</td>
<td width="16%" align="center" class="td">会员名字</td>
<td width="14%" align="center" class="td">登陆次数</td>
<td width="14%" align="center" class="td">是否审核</td>
<td width="17%" align="center" class="td">注册时间</td>
<td width="7%" align="center" class="td">修改</td>
<td width="7%" align="center" class="td">删除</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?del=checkbox"> 
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from [user] order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无会员!</div>")
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
for i=1 to rs.pagesize  %> 
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="5%" align="center" class="td"><input type="text" readonly="readonly" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="17%" height="25" align="center" class="td"><a href="xiugai_user.asp?id=<%=rs("id")%>" style="color:#003399"><%=rs("useradmin")%></a> </td>
<td width="16%" align="center" class="td"><%=rs("zsname")%></td>
<td width="14%" align="center" class="td">登陆次数：<%=rs("dlcs")%></td>
<td width="14%" align="center" class="td"><%if rs("sh")=1 then
response.Write("已审核")
else
response.Write("<font color=#FF0000>未审核</font>")
end if%></td>
<td width="17%" align="center" class="td"><%=rs("data")%></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit3" value="修改" onclick="window.location.href='admin_user.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="3%" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="5%" align="center" class="td">全选</td>
<td height="25" colspan="7" class="td">
<input type="submit" class="btn" onclick="form.action='?sh=ok';" value="批量审核"/>
<input type="submit" class="btn" onclick="form.action='?sh=no';" value="取消审核"/>
<input type="submit" class="btn" onclick="form.action='?del=ok';" value="批量删除"/>
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
<% End If %>
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [user] where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<form  name="add" method="post" action="admin_user.asp?action=xiugai&id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">查看/修改会员资料</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="13%" height="28" align="right" class="td">账号ID</td>
        <td width="87%"  class="td"><%=rs("id")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr>
        <td width="13%" height="25" align="right" class="td">会员帐号</td>
        <td class="td"><%=rs("useradmin")%></td>
      </tr>
      <tr>
        <td width="13%" height="25" align="right" class="td">会员密码</td>
        <td class="td"><input name="userpassword" type="text" size="30"  /> 不修改请留空! </td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">真实姓名</td>
        <td class="td"><%=rs("zsname")%></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">会员性别</td>
        <td class="td"><input type="radio" name="sex" value="0" <%if rs("sex")=0 then%>checked<%end if%>>先生　 
<input type="radio" name="sex" value="1" <%if rs("sex")=1 then%>checked<%end if%>>女士</td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">密码保护问题</td>
        <td class="td"><%=rs("wen")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">密码保护答案</td>
      <td class="td"><%=rs("da")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">公司名称</td>
      <td class="td"><label>
        <input name="gsname" type="text" value="<%=rs("gsname")%>" size="40" />
      </label></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">公司地址</td>
      <td class="td"><input name="gsadd" type="text" value="<%=rs("gsadd")%>" size="40" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">邮政编码</td>
      <td class="td"><input name="youbian" type="text" value="<%=rs("youbian")%>" size="20" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">联系电话</td>
      <td class="td"><input name="tel" type="text" value="<%=rs("tel")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">联系传真</td>
      <td class="td"><input name="fax" type="text" value="<%=rs("fax")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">联系手机</td>
      <td class="td"><input name="sj" type="text" value="<%=rs("sj")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">电子邮箱</td>
      <td class="td"><input name="mail" type="text" value="<%=rs("mail")%>" size="30" /></td>
    </tr>
  <tr>
      <td height="25" align="right" class="td">公司网址</td>
      <td class="td"><input name="wz" type="text" value="<%=rs("wz")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">注册时间</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
  <tr>
      <td height="25" align="right" class="td">最后登陆</td>
      <td class="td"><%=rs("dldata")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">登陆次数</td>
      <td class="td"><%=rs("dlcs")%> 次</td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">注册IP</td>
      <td class="td"><%=rs("ip")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">&nbsp;</td>
      <td class="td"><input type="submit" name="button" id="button" value="更新资料" class="btn"/></td>
    </tr>
    
    </table>
    </td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from [user] where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('删除成功！');window.location.href='admin_user.asp';</script>"
end if 
%>
<%
if Request.QueryString("sh")="ok" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_user.asp';</script>" 
response.end()
end if
sql="update [user] set sh=1 where id in ("&Request("id")&")" 
conn.execute(sql)
Response.Write "<script>alert('恭喜!审核成功!');window.location.href='admin_user.asp';</script>" 
end if
 
if Request.QueryString("sh")="no" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_user.asp';</script>" 
response.end()
end if
sql="update [user] set sh=0 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('恭喜!取消审核成功!');window.location.href='admin_user.asp';</script>" 
end if

if Request.QueryString("del")="ok" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_user.asp';</script>" 
response.end()
end if
sql="delete from [user] where id in ("&Request("id")&")"
conn.Execute ( sql )
Response.Write "<script>alert('恭喜!操作成功!');window.location.href='admin_user.asp';</script>" 
end if

'=====================
if Request.QueryString("xiugai")="ok" then 
id=request("id")
userpassword=request.form("userpassword")
sex=request.form("sex")
gsname=request.form("gsname")
gsadd=request.form("gsadd")
youbian=request.form("youbian")
tel=request.form("tel")
fax=request.form("fax")
sj=request.form("sj")
mail=request.form("mail")
wz=request.form("wz")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from [user] where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
if request.form("userpassword") <> "" then
rs("userpassword")=md5(request.form("userpassword"))
end if
rs("sex")=sex
rs("gsname")=gsname
rs("gsadd")=gsadd
rs("youbian")=youbian
rs("tel")=tel
rs("fax")=fax
rs("sj")=sj
rs("mail")=mail
rs("wz")=wz
rs.update 
rs.close 
response.write "<script>alert('资料修改成功！');window.location.href='admin_user.asp?action=admin';</script>" 
end if%>