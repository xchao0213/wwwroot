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
exec="select * from link where id="& id 
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
<title>无标题文档</title>
<link rel="stylesheet" href="../zycheditor/themes/default/default.css" />
<link rel="stylesheet" href="../zycheditor/plugins/code/prettify.css" />
<script charset="utf-8" src="../zycheditor/kindeditor.js"></script>
<script charset="utf-8" src="../zycheditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zycheditor/plugins/code/prettify.js"></script>
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
								K('#logo').val(url);
								editor.hideDialog();
							}
						});
					});
				});
			});
		</script>
</head>

<body>
<form  name="myform" method="post" action="xiugai_link_img.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改图片链接</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8" >
        <td width="15%" height="12" class="td">链接名称</td>
        <td width="85%"  class="td"><input name="title" type="text" value="<%=rs("title")%>" size="30"  /></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" width="15%" class="td">链接地址</td>
        <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="30"  /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="13" class="td">LOGO图标</td>
        <td class="td"><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="32%"><input name="logo" type="text" value="<%=rs("logo")%>" size="30" /></td>
    <td width="11%"><input name="button" type="button" id="filemanager" value="选取已上传图片" class="btn" style="height:20px;" /></td>
    <td width="57%"><iframe src="up.asp?formname=myform&inputname=logo&uploadstyle=Logo" frameborder="0" scrolling="no" width="350" height="21"></iframe></td>
  </tr>
</table></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="6" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="30"  /></td>
      </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/>
		<input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
    </table></td>
  </tr>
</table></form>
</body>
</html>
<% 
if request("xiugai")="ok" then
id=request("id")
title=request.form("title")
url=request.form("url")
logo=request.form("logo")
px_id=request.form("px_id")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from link where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
	if title=""  then 
response.Write("<script language=javascript>alert('链接名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("title")=title
rs("url")=url
rs("logo")=logo
rs("px_id")=px_id
rs.update 
rs.close 
response.write "<script>alert('修改成功！');window.location.href='admin_link_img.asp';</script>"
end if 
%>