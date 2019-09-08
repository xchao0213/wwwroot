<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>管理员管理</title>
<SCRIPT language=javascript>
function SelectAll(form){
  for (var i=0;i<form.AdminMight.length;i++){
    var e = form.AdminMight[i];
    if (e.disabled==false)
       e.checked = form.chkAll.checked;
    }
}
</script>
</head>
<body>
<%if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">管理员管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="8%" align="center" class="td">ID</td>
<td width="16%" height="25" align="center" class="td">管理员</td>
<td width="18%" align="center" class="td">权 限</td>
<td width="15%" align="center" class="td">登陆次数</td>
<td width="26%" align="center" class="td">最后登陆</td>
<td width="7%" align="center" class="td">修改 </td>
<td width="10%" align="center" class="td">删除</td>
</tr></thead>
<%	
set rs=server.createobject("adodb.recordset")
if key=0 then
exec="select * from admin order by id asc" 
else
exec="select * from admin where id="&session("admin")&"" 
end if
rs.open exec,conn,1,1 
if rs.eof then
response.Write "暂无信息！"
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
for i=1 to rs.pagesize
if rs("key")<>0 then
key="<font color=#336699>普通管理员</font>"
else
key="<font color=#336699>超级管理员</font>"
end if%>
<tr>
<td width="8%" align="center" class="td"><%=rs("id")%></td>
<td width="16%" height="25" align="center" class="td"><%=rs("admin")%> </td>
<td width="18%" align="center" class="td"><%=key%></td>
<td width="15%" align="center" class="td"><%=rs("dlcs")%></td>
<td width="26%" align="center" class="td"><%=rs("dldata")%></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit3" value="修改" onClick="window.location.href='admin_administrator.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/> </td>
<td width="10%" align="center" class="td"><input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_administrator.asp?action=admin&id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<% rs.movenext 
if rs.eof then exit for 
next 
end if%></table>
</td>
  </tr>
</table>
<div style="margin-top:10px">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?add=ok" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加管理员</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="12%" height="25" align="right" class="td">管理员帐号</td>
        <td width="88%"  class="td"><input name="admin" type="text" size="30"  /> * (请勿使用中文!)</td>
      </tr>
      <tr>
        <td width="12%" height="13" align="right" class="td">登陆密码</td>
        <td class="td"><input name="password" type="text" size="30"  />
          * </td>
      </tr>
      <tr>
        <td width="12%" height="12" align="right" class="td">确认密码</td>
        <td class="td"><input name="password3" type="text" size="30"  />
          * </td>
      </tr>
      <tr>
        <td width="12%" height="25" align="right" class="td">真实姓名</td>
        <td class="td"><input name="zsname" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">管理员QQ</td>
        <td class="td"><input name="qq" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">管理员级别</td>
        <td class="td"><label>
          <select name="key" id="key">
            <option value="1">普通管理员</option>
            <option value="0">超级管理员</option>
          </select></label></td>
      </tr>
      <tr>
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="提交数据"  class="btn"/></td>
      </tr>
    </table></td>
  </tr></form>
</table>
</div>
<%end if%>
<!--以下为超级管理员修改资料-->
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from admin where id="&id
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 %> 
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="admin_administrator.asp?action=xiugai<%=rs("id")%>&adminxiugai=ok" method="post" name="add">
   <tr>
     <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改管理员资料</div></td>
   </tr>
   <tr>
     <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
         <tr >
           <td width="12%" height="25" align="right" class="td">管理员帐号</td>
           <td width="88%"  class="td"><input name="id" type="hidden" value="<%=rs("id")%>" size="30"  />
		   <input name="admin" type="text" value="<%=rs("admin")%>" size="30"  /></td>
        </tr>
         <tr>
           <td width="12%" align="right" class="td">登陆密码</td>
           <td class="td"><input name="password" type="text" size="30"  /> 不修改请留空！</td>
         </tr>
            <tr>
           <td width="12%" align="right" class="td">确认密码</td>
           <td class="td"><input name="password2" type="text" size="30"  /></td>
         </tr>
        <tr>
           <td width="12%" height="25" align="right" class="td">真实姓名</td>
           <td class="td"><input name="zsname" type="text" value="<%=rs("zsname")%>" size="30"  /></td>
         </tr>
              <tr>
           <td height="25" align="right" class="td">管理员QQ</td>
           <td class="td"><input name="qq" type="text" value="<%=rs("qq")%>" size="30"  /></td>
         </tr>
         <tr>
           <td height="25" align="right" class="td">管理员级别</td>
           <td class="td">
<%if key=0 then%><input type="radio" name="key" value="0" <%if rs("key")=0 then%>checked<%end if%> onClick="Date3.style.display='none'">超级管理员<%end if%> 
<input type="radio" name="key" value="1" <%if rs("key")=1 then%>checked<%end if%> onClick="Date3.style.display=''">普通管理员　 
 </td>
         </tr>
<%'if key=0 then%>
<tr id="Date3" style="display:<%If rs("key")=0 Then Response.Write "none"%>">
<td height="25" colspan="2"  class="td">
<% dim MightList(5,10)
MightList(0,0) = "网站设置"
MightList(0,1) = "全局参数-1"
MightList(0,2) = "附加参数-2"
MightList(0,3) = "导航菜单-3" 
MightList(0,4) = "管理员管理-4"
MightList(0,5) = "登陆记录-5"
MightList(0,6) = "SQL管理-6"
MightList(0,7) = "生成静态-7"
MightList(0,8) = "空间占用查看-8"

MightList(1,0)="内容管理1"
MightList(1,1)="单页管理-9"
MightList(1,2)="新闻管理-10"
MightList(1,3)="新闻分类-10"
MightList(1,4)="案例管理-12"
MightList(1,5)="案例分类-13"
MightList(1,6)="产品管理-14"
MightList(1,7)="产品分类-15"
MightList(1,8)="团队管理-16"
MightList(1,9)="服务管理-17"
MightList(1,10)="视频管理-18"

MightList(2,0)="内容管理2"
MightList(2,1)="视频分类-19"
MightList(2,2)="下载管理-20"
MightList(2,3)="下载分类-21"
MightList(2,4)="相册管理-22"
MightList(2,5)="相册分类-23"
MightList(2,6)="人才管理-24"

MightList(3,0)="辅助功能"
MightList(3,1)="JS广告-25"
MightList(3,2)="首页大图-26"
MightList(3,3)="频道幻灯-27"
MightList(3,4)="会员管理-28"
MightList(3,5)="在线客服-29"
MightList(3,6)="订单管理-30"
MightList(3,7)="数据备份-31"
MightList(3,8)="附件管理-32"

MightList(4,0)="其它管理"
MightList(4,1)="留言管理-33"
MightList(4,2)="地图生成-34"
MightList(4,3)="友情链接-35"
MightList(4,4)="网站推广-36"
MightList(4,5)="SEO管理-37"%>   
<fieldset style="margin:2px 2px 2px 2px">
<legend><B>系统权限内容</B>	<label><input name="chkAll" type="checkbox" id="chkAll" onClick="SelectAll(this.form)">全选</label></legend>
<%
Dim n,i
for i=0 to ubound(MightList,1)
If i<>0 then
Response.Write "<br><br>"
End if
Response.Write "<b>"&MightList(i,0)&"</b><br>"
for	j=1	to ubound(MightList,2)
if isempty(MightList(i,j)) then exit for
tmpmenu=split(MightList(i,j),"-")
menuname=tmpmenu(0)
menurl=tmpmenu(1)%>
    <label class="box"><input name="manage" type="checkbox" id="AdminMight" value="<%=menurl%>"<% if instr(","& rs("manage") &",",","&menurl&",")>0 then response.write " checked" %>><%=menurl%>.<%=menuname%></label>
    <%next
next%>
</fieldset></td>
</tr>
<%' End If %>
         <tr>
           <td height="25" align="right" class="td">&nbsp;</td>
           <td class="td"><input type="submit" name="button2" id="button2" value="确定修改"  class="btn"/></td>
         </tr>
     </table></td>
   </tr></form>
 </table>
<%end if%>
</body>
</html>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,3
admin=request.form("admin")
password=request.form("password")
password3=request.form("password3")
zsname=request.form("zsname")
qq=request.form("qq")
key=request.form("key")
if admin=""  then 
response.Write("<script language=javascript>alert('登陆帐号不能为空!');history.go(-1)</script>") 
response.end 
end if
letters="0123456789abcdefghijklmnopqrstuvwxyz" 
admin=Lcase(trim(Request.Form("admin"))) 
for i=1 to len(admin) 
u=mid(admin,i,1) 
if Instr(letters,u)=0 then 
response.write "<script>alert('登陆帐号只能由字母、数字及下划线组成!');history.go(-1);</script>" 
response.end 
end if 
next 
if password=""  then 
response.Write("<script language=javascript>alert('登陆密码不能为空!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
if request.Form("password")<>request.Form("password3") then 
Response.Write "<script>alert('密码输入不一致！请重新输入！');history.go(-1);</script>" 
response.end
end if
rs("admin")=admin
rs("password")=md5(request.form("password"))
rs("zsname")=zsname
rs("qq")=qq
rs("key")=key
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('添加成功！');window.location.href='admin_administrator.asp?action=admin';</script>" 
end if

if request("adminxiugai")="ok" then
id=request("id")
if id="" or not isnumeric(id) then
Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('admin_administrator.asp?action=admin');</script>"
Response.End()
end if
SQL="Select * from admin where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script language='javascript'>alert('参数不正确!');document.location.href('admin_administrator.asp?action=admin');</script>"
Response.End()
end if
if request.form("password")<>"" then
rs("password")=md5(request.form("password"))
else
end if
if request.Form("password")<>request.Form("password2") then 
Response.Write "<script>alert('密码输入不一致！请重新输入！');history.go(-1);</script>" 
response.end
end if
rs("admin")=request.form("admin")
rs("zsname")=request.form("zsname")
rs("qq")=request.form("qq")
if key=0 then
rs("manage")=replace(request.form("manage")," ","")
rs("key")=request.form("key")
end if
rs.update 
rs.close 
response.write "<script>alert('管理员资料修改成功！');window.location.href='admin_administrator.asp?action=admin';</script>" 
end if

if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from admin where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('删除成功！');window.location.href='admin_administrator.asp?action=admin';</script>"
end if %>