<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%
if session("qx")=2 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px; "")>您没有管理该模块的权限！</font></div>")
response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>增加团队</title>
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
</script>
</head>
<body>
<form  name="add" method="post" action="add_tuandui.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加团队成员</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
<tr >
  <td width="10%" height="28" align="right" class="td">姓 名 <font color="#FF0000">*</font></td>
  <td  class="td">
    <input name="name" type="text" size="40"  />
    <label></label></td>
  </tr>
<tr>
  <td width="10%" height="13" align="right" class="td">职 务 <font color="#FF0000">*</font></td>
  <td class="td"><input name="zw" type="text" size="40" /></td>
</tr>
 <tr>
  <td height="13" align="right" class="td">照 片 <font color="#FF0000">*</font></td>
  <td class="td"><table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><input type="text" id="url" name="img" size="40" /></td>
      <td><input type="button" class="btn2" id="filemanager" value="选取已上传图片" /></td>
      <td><iframe src="up.asp?formname=add&inputname=img&uploadstyle=tuandui" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
    </tr>
  </table></td>
  </tr>

 <tr>
   <td height="25" align="right" class="td">排 序 <font color="#FF0000">*</font></td>
   <td class="td"><label>
     <input name="px_id" type="text" value="1"  size="20" />
   </label></td>
   </tr>
  <tr>
<td height="25" align="right" class="td">个人博客</td>
<td class="td"><input name="url" type="text" value="#" size="60" /></td>
  </tr>
  <tr>
<td height="25" align="right" class="td">专项<font color="#FF0000">*</font></td>
<td class="td"><input name="unit" type="text" value="" size="60" /></td>
  </tr>
<tr>
  <td height="25" align="right" class="td">个人简介 <font color="#FF0000">*</font></td>
  <td class="td"><textarea name="body" style="width:700px;height:200px;visibility:hidden;"></textarea></td>
</tr>
<tr>
   <td height="12" align="right" class="td">&nbsp;</td>
   <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
 </tr>
</table>
</td>
  </tr>
</table></form>
</body>
</html>
<%
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
rs("body")=body
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('团队增加成功，点击继续增加！');window.location.href='admin_tuandui.asp';</script>" 
end if
%>