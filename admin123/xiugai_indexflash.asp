<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%
if session("qx")=1 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px;"">您没有管理该模块的权限！</font></div>")
response.End
End If
%>
<% 
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from indexflash where id="& id 
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
<title>修改幻灯</title>
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
<form  name="myform" method="post" action="xiugai_indexflash.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改幻灯</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8" >
        <td height="28" width="16%" class="td">名称</td>
        <td  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="30"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="30"  /></td>
      </tr>
      <tr bgcolor="#ffffff">
        <td height="25" width="16%" class="td">图片地址 <font color="#FF0000">*</font></td>
        <td colspan="2" class="td"><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%"><input type="text" name="img" id="url" size="30" value="<%=rs("img")%>"></td>
    <td width="11%"><input name="button2" type="button" id="filemanager" value="选取已上传图片" class="btn" style="height:20px;" /></td>
    <td width="59%"><iframe src="up.asp?formname=myform&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="340" height="21"></iframe></td>
  </tr>
</table>          </td>
        </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" width="16%" class="td">链接地址</td>
        <td class="td"><input name="link" type="text" value="<%=rs("link")%>" size="30"  /> 
          空链接请用#号</td>
      </tr>
	  
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="10"  /></td>
      </tr>
	<!--
	 <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" width="16%" class="td">幻灯动画方式</td>
          <td class="td"><select name="slicing">
          <option value=""<%if rs("slicing")="" then%>selected<%end if%>>水平或垂直方向</option>
          <option value="horizontal" <%if rs("slicing")="horizontal" then%>selected<%end if%>>horizontal水平</option>
          <option value="vertical"  <%if rs("slicing")="vertical" then%>selected<%end if%>>vertical垂直</option>
        </select>
          <label>
		<select name="direction">
		 <option value="" <%if rs("direction")="" then%>selected<%end if%>>选择方向</option>
          <option value="left" <%if rs("direction")="left" then%>selected<%end if%>>left向左</option>
          <option value="right"<%if rs("direction")="right" then%>selected<%end if%>>right向右</option>
          <option value="down" <%if rs("direction")="down" then%>selected<%end if%>>down向下</option>
          <option value="up" <%if rs("direction")="up" then%>selected<%end if%>>up向上</option>
        </select>
          </label>
          <label>		
		<select name="num">
          <option value="" <%if rs("num")="" then%>selected<%end if%>>切片数量</option>
          <option value="1" <%if rs("num")="1" then%>selected<%end if%>>1片</option>
		  <option value="2" <%if rs("num")="2" then%>selected<%end if%>>2片</option>
          <option value="3" <%if rs("num")="3" then%>selected<%end if%>>3片</option>
          <option value="4" <%if rs("num")="4" then%>selected<%end if%>>4片</option>
          <option value="5" <%if rs("num")="5" then%>selected<%end if%>>5片</option>
		  <option value="6" <%if rs("num")="6" then%>selected<%end if%>>6片</option>
        </select>
          </label>
          <label>
		<select name="shader">
          <option value="" <%if rs("shader")="" then%>selected<%end if%>>无切片阴影</option>
          <option value="phong" <%if rs("shader")="phong" then%>selected<%end if%>>有切片阴影</option>
        </select>
          </label>
          <label>
		<select name="delay">
          <option value="" <%if rs("delay")="" then%>selected<%end if%>>延迟时间</option>
          <option value="0.03" <%if rs("delay")="0.03" then%>selected<%end if%>>延迟0.03</option>
          <option value="0.05" <%if rs("delay")="0.05" then%>selected<%end if%>>延迟0.05</option>
          <option value="0.1" <%if rs("delay")="0.1" then%>selected<%end if%>>延迟0.1</option>
        </select>
          </label>
            </td>
        </tr>-->
	<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">背景颜色 <font color="#FF0000">*</font></td>
        <td class="td"><input name="coler" type="text" style="background:<%=rs("coler")%>; color:#FFFFFF" value="<%=rs("coler")%>"  size="10"  /> 
        宽屏背景色</td>
      </tr>	 
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="确认修改"  class="btn"/></td>
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
title=request.form("title")
img=request.form("img")
link=request.form("link")
px_id=request.form("px_id")
direction=request.form("direction")
slicing=request.form("slicing")
num=request.form("num")
shader=request.form("shader")
delay=request.form("delay")
coler=request.form("coler")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from indexflash where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
if title=""  then 
response.Write("<script language=javascript>alert('广告标题不能为空!');history.go(-1)</script>") 
response.end 
end if
if link=""  then 
response.Write("<script language=javascript>alert('链接地址不能为空!');history.go(-1)</script>") 
response.end 
end if
if img="" then 
response.Write("<script language=javascript>alert('图片地址不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id="" then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("title")=title
rs("img")=img
rs("link")=link
rs("px_id")=px_id
rs("direction")=direction
rs("slicing")=slicing
rs("num")=num
rs("shader")=shader
rs("delay")=delay
rs("coler")=coler
rs.update 
rs.close 
response.write "<script>alert('主页幻灯修改成功！');window.location.href='admin_indexflash.asp';</script>" 
end if
%> 
