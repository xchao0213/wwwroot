<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%
if session("qx")=1 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px; "")>您没有管理该模块的权限！</font></div>")
response.End
End If 
exec="select * from tuandui where id="& request.QueryString("id") 
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
<form  name="add" method="post" action="xiugai_tuandui.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改团队成员</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
  <tr >
    <td width="9%" height="28" align="right" class="td">姓 名 <font color="#FF0000">*</font></td>
    <td width="91%"  class="td">
      <input name="name" type="text" value="<%=rs("name")%>" size="40"  />
      <label>
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>" />
      </label></td>
    </tr>
  <tr>
    <td width="9%" height="13" align="right" class="td">职 务 <font color="#FF0000">*</font></td>
    <td class="td"><input name="zw" type="text" value="<%=rs("zw")%>" size="40" /></td>
  </tr>
   <tr>
    <td height="13" align="right" class="td">照 片 <font color="#FF0000">*</font></td>
    <td class="td"><table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
          <td><input name=img id=url type=text value="<%=rs("img")%>" size=50 /></td>
          <td><input name="button2" type="button" class="btn2" id="filemanager" value="选取已上传图片" /></td>
          <td><iframe src="up.asp?formname=add&inputname=img&uploadstyle=tuandui" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
        </tr>
  </table></td>
    </tr>
   <tr>
     <td height="25" align="right" class="td">排 序 <font color="#FF0000">*</font></td>
     <td class="td"><label>
       <input name="px_id" type="text" value="<%=rs("px_id")%>"  size="13" />
     </label></td>
     </tr>
    <tr>
  <td height="25" align="right" class="td">个人博客</td>
  <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="60" /></td>
    </tr>
    <tr>
  <td height="25" align="right" class="td">专项<font color="#FF0000">*</font></td>
  <td class="td"><input name="unit" type="text" value="<%=rs("unit")%>" size="60" /></td>
    </tr>
 <tr>
    <td height="25" align="right" class="td">个人简介 <font color="#FF0000">*</font></td>
    <td class="td"><textarea name="body" style="width:700px;height:200px;"><%=rs("body")%></textarea></td>
 </tr>
 <tr>
   <td height="12" align="right" class="td">显示位置</td>
   <td class="td"><label for="top"><input name="top" type="checkbox" id="top" value="1" <% If rs("top")=1 Then %>checked<% End If %> />置顶</label> 
                  <label for="pic"><input name="pic" type="checkbox" id="pic" value="1" <% If rs("pic")=1 Then %>checked<% End If %>/>照片</label></td>
 </tr>
 <tr>
     <td height="12" class="td">&nbsp;</td>
     <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
   </tr>
</table>
</td>
</tr>
</table></form>
</body>
</html>
<% 
if Request.QueryString("xiugai")="ok" then 
id=request("id")
name=request.form("name")
zw=request.form("zw")
img=request.form("img")
url=request.form("url")
px_id=request.form("px_id")
unit=request.form("unit")
top=request.form("top")
pic=request.form("pic")
body=request.form("body")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from tuandui where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if
if name=""  then 
response.Write("<script language=javascript>alert('姓名不能为空!');history.go(-1)</script>") 
response.end 
end if
if zw=""  then 
response.Write("<script language=javascript>alert('职务不能为空!');history.go(-1)</script>") 
response.end 
end if
rs("name")=name
rs("zw")=zw
rs("img")=img
rs("url")=url
rs("px_id")=px_id
rs("unit")=unit
if top<>"" then rs("top")=top else rs("top")=0
if pic<>"" then rs("pic")=pic else rs("pic")=0
rs("body")=body
rs.update 
rs.close 
response.write "<script>alert('团队成员修改成功！');window.location.href='admin_tuandui.asp?action=admin';</script>" 
end if
%> 
