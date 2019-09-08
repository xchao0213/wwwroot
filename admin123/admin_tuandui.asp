<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(16)
id=request.QueryString("id")
key=request.QueryString("key")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>管理团队成员</title>
<link rel="stylesheet" href="../zycheditor/themes/default/default.css" />
<link rel="stylesheet" href="../zycheditor/plugins/code/prettify.css" />
<script charset="utf-8" src="../zycheditor/kindeditor.js"></script>
<script charset="utf-8" src="../zycheditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zycheditor/plugins/code/prettify.js"></script>
<script>
KindEditor.ready(function(K) {
  var editor1 = K.create('textarea[name="body"]', {
      cssPath : '../zycheditor/plugins/code/prettify.css',
      uploadJson : '../zycheditor/asp/upload_json.asp',
      fileManagerJson : '../zycheditor/asp/file_manager_json.asp',
      allowFileManager : true,
      afterCreate : function() {
          var self = this;
          K.ctrl(document, 13, function() {
              self.sync();
              K('form[name=add]')[0].submit();
          });
          K.ctrl(self.edit.doc, 13, function() {
              self.sync();
              K('form[name=add]')[0].submit();
          });
      }
  });
  prettyPrint();
});
</script>
<script>
KindEditor.ready(function(K) {
var editor = K.editor({
	fileManagerJson : '../zycheditor/asp/file_manager_json.asp'
});
K('#filemanager').click(function() {
	editor.loadPlugin('filemanager', function() {
		editor.plugin.filemanagerDialog({
			viewType : 'VIEW',
			dirName : 'image',
			clickFn : function(url, title) {
				K('#url').val(url);
				editor.hideDialog();
			}
		});
	});
});
});

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
<%if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">管理团队成员</div></td>
  </tr>
  <tr>
     <td colspan="2">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="50" align="center" class="td"><input type="checkbox" /></td>
<td width="50" align="center" class="td">ID</td>
<td width="130" align="center" class="td">照片</td>
<td class="td">姓 名</td>
<td width="100" align="center" class="td">职 务</td>
<td width="60" align="center" class="td">博客地址</td>
<td width="60" align="center" class="td">排 序</td>
<td width="120" align="center" class="td">推 荐</td>
<td width="60" align="center" class="td">修 改</td>
<td width="60" align="center" class="td">删 除</td>
</tr></thead>
<form id="form" name="form" method="post" action="?del=checkbox"> 
<%set rs=server.createobject("adodb.recordset") 
exec="select * from tuandui order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无成员!</div>")
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
for i=1 to rs.pagesize %>
<tr>
<td height="38" align="center" class="td"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>" /></td>
<td align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td align="center" class="td"><a href="<%=rs("img")%>" target="_blank"><img src="<%=rs("img")%>" alt="点击查看大图" width="100" height="60" border="0" /></a></td>
<td class="td"><%=rs("name")%></td>
<td align="center" class="td"><%=rs("zw")%></td>
<td align="center" class="td">[<a href="http://<%=rs("url")%>" target="_blank" style="color:#003399">访问</a>]</td>
<td align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td align="center" class="td"><% If rs("top")=1 Then %>置顶<% End If %>---<% If rs("pic")=1 Then %>照片<% End If %></td>
<td align="center" class="td"><input type="button" name="Submit3" value="修改" onclick="window.location.href='xiugai_tuandui.asp?id=<%=rs("id")%>' "  class="btn"/></td>
<td align="center" class="td"><input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_tuandui.asp?act=del&id=<%=rs("id")%>';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="20" height="30" align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="40" height="30" align="center">全选</td>
<td height="30" colspan="8"><label>
  <input type="submit" class="btn" onclick="form.action='?action=del';" value="删除"/>
  <input type="submit" class="btn" onclick="form.action='?action=ordsc';" value="更新排序"/>
</label><%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>
</td>
      </tr>
    </table>
<%end if%>
<%if request.querystring("action")="add" then%>
<form  name="add" method="post" action="admin_tuandui.asp?action=add&add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加团队成员</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
<tr >
  <td width="10%" height="28" align="right" class="td"><font color="#FF0000">*</font>姓 名</td>
  <td  class="td"><input name="name" type="text" size="40"  /></td>
  </tr>
<tr>
  <td width="10%" height="13" align="right" class="td"><font color="#FF0000">*</font>职 务</td>
  <td class="td"><input name="zw" type="text" size="40" /></td>
</tr>
 <tr>
  <td height="13" align="right" class="td"><font color="#FF0000">*</font>照 片</td>
  <td class="td"><table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><input type="text" id="url" name="img" size="40" /></td>
      <td><input type="button" class="btn2" id="filemanager" value="选取已上传图片" /></td>
      <td><iframe src="up.asp?formname=add&inputname=img&uploadstyle=tuandui" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
    </tr>
  </table></td>
  </tr>

 <tr>
   <td height="25" align="right" class="td"><font color="#FF0000">*</font>排 序</td>
   <td class="td"><label>
     <input name="px_id" type="text" value="1"  size="20" />
   </label></td>
   </tr>
  <tr>
<td height="25" align="right" class="td">个人博客</td>
<td class="td"><input name="url" type="text" value="#" size="60" /></td>
  </tr>
  <tr>
<td height="25" align="right" class="td"><font color="#FF0000">*</font>专项</td>
<td class="td"><input name="unit" type="text" value="" size="60" /></td>
  </tr>
<tr>
  <td height="25" align="right" class="td"><font color="#FF0000">*</font>个人简介</td>
  <td class="td"><textarea name="body" style="width:700px;height:200px;visibility:hidden;"></textarea></td>
</tr>
<tr>
   <td height="12" align="right" class="td">显示位置</td>
   <td class="td"><label for="top"><input name="top" type="checkbox" id="top" value="1" /> 置顶</label> 
     <label for="pic"><input type="checkbox" name="pic" id="pic" value="1" />照片</label></td>
 </tr>
<tr>
   <td height="12" align="right" class="td">&nbsp;</td>
   <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
 </tr>
</table>
</td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%if request("act")="del" then
	id=request("ID")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('admin_tuandui.asp?action=admin');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from tuandui where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');document.location.href('admin_tuandui.asp?action=admin');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('当前团队成员删除成功！');document.location.href('admin_tuandui.asp?action=admin');</script>"
end if
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [tuandui] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_tuandui.asp?action=admin';</script>" 
end if

if Request.QueryString("action")="ordsc" then
if request.form("px_id")="" then
Response.Write "<script language='javascript'>alert('排序ID不能为空');document.location.href('admin_tuandui.asp?action=admin');</script>"
Response.End()
end if
depname=request.form("px_id")
myarray=Split(depname,", ") 
set rs=server.createobject("adodb.recordset")
sql="select * from tuandui order by px_id asc"
rs.open sql,conn,1,3
for i= 0 to Ubound(myarray)
rs("px_id")=myarray(i)
rs.update
rs.movenext
next
rs.close
set rs=nothing
Response.Write "<script>alert('操作成功!');window.location.href='admin_tuandui.asp?action=admin';</script>" 
end if

if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from tuandui"
rs.open sql,conn,1,3
name=request.form("name")
zw=request.form("zw")
img=request.form("img")
url=request.form("url")
px_id=request.form("px_id")
unit=request.form("unit")
ssfl=request.form("ssfl")
body=request.form("body")
top=request.form("top")
pic=request.form("pic")
if name=""  then 
response.Write("<script language=javascript>alert('姓名不能为空!');history.go(-1)</script>") 
response.end 
end if
if zw=""  then 
response.Write("<script language=javascript>alert('职务不能为空!');history.go(-1)</script>") 
response.end 
end if
if img=""  then 
response.Write("<script language=javascript>alert('照片不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排列顺序不能为空!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
rs("name")=name
rs("zw")=zw
rs("img")=img
rs("px_id")=px_id
rs("unit")=unit
if top<>"" then rs("top")=top else rs("top")=0
if pic<>"" then rs("pic")=pic else rs("pic")=0
rs("body")=body
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('团队增加成功，点击继续增加！');window.location.href='admin_tuandui.asp?action=admin';</script>" 
end if
%>