<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(20)%>
<% 
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from download where id="& id 
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
<title>修改下载</title>
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
<form  name="myform" method="post" action="xiugai_download.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改下载</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="11%" height="28" align="right" class="td">下载名称</td>
        <td width="89%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="50"  />
          <input name="id" type="hidden" value="<%=rs("id")%>"  /></td>
      </tr>
      <tr>
        <td width="11%" height="25" align="right" class="td">所属分类</td>
        <td class="td">
<select name="ssfl" id="select">
<%set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from download_fl",conn,1,1
while not rsc.eof
if rs("ssfl")=rsc("id") then
response.Write("<option value=""" & rsc("id") & """ selected>" & rsc("title") & "</option>")
else
response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
%>
</select> 下载权限：<select name="qx" id="select2">
          <option value="1" <%if rs("qx")=1 then%>selected<%end if%>>所有游客</option>
          <option value="9" <%if rs("qx")=9 then%>selected<%end if%>>本站会员</option>
        </select></td>
      </tr>
      <tr>
        <td width="11%" height="7" align="right" class="td">程序语言</td>
        <td class="td">
<input type="radio" name="yuyan" value="简体中文" <%if rs("yuyan")="简体中文" then%>checked<%end if%>>简体中文　 
<input type="radio" name="yuyan" value="繁体中文　" <%if rs("yuyan")="繁体中文" then%>checked<%end if%>>繁体中文　 
<input type="radio" name="yuyan" value="English" <%if rs("yuyan")="English" then%>checked<%end if%>>English
<input type="radio" name="yuyan" value="多国语言" <%if rs("yuyan")="多国语言" then%>checked<%end if%>>多国语言 </td>
      </tr>
      <tr>
        <td width="11%" height="6" align="right" class="td">运行平台</td>
        <td class="td">
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"XP",0)>0 then response.write" checked" end if%> value="XP"/>XP </label>
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"Vista",0)>0 then response.write" checked" end if%> value="Vista" />Vista</label>
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"Linux",0)>0 then response.write" checked" end if%> value="Vista" />Linux</label>
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"Win7",0)>0 then response.write" checked" end if%> value="Win7" checked="checked" />Win7</label>
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"Win8",0)>0 then response.write" checked" end if%> value="Win8" />Win8</label>
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"android",0)>0 then response.write" checked" end if%> value="android" />android</label>
<label><input name="yxpt" type="checkbox" <%if InStr(1,rs("yxpt"),"IOS",0)>0 then response.write" checked" end if%> value="IOS" />IOS</label>
</td>
      </tr>
      <tr>
        <td width="11%" height="12" align="right" class="td">推荐等级</td>
        <td class="td"><label>
          <select name="tjdj" id="select2">
            <option value="★☆☆☆☆" <%if rs("tjdj")="★☆☆☆☆" then%>selected<%end if%>>★☆☆☆☆</option>
            <option value="★★☆☆☆" <%if rs("tjdj")="★★☆☆☆" then%>selected<%end if%>>★★☆☆☆</option>
            <option value="★★★☆☆" <%if rs("tjdj")="★★★☆☆" then%>selected<%end if%>>★★★☆☆</option>
            <option value="★★★☆☆" <%if rs("tjdj")="★★★☆☆" then%>selected<%end if%>>★★★★☆</option>
            <option value="★★★★★" <%if rs("tjdj")="★★★★★" then%>selected<%end if%>>★★★★★</option>
          </select>
        </label></td>
      </tr>
    <tr>
        <td height="13" align="right" class="td">程序大小</td>
        <td class="td"><input name="daxiao" type="text" value="<%=rs("daxiao")%>" size="10"  />
<label><input name="danwei" type="radio" value="KB" <%if rs("danwei")="KB" then%>checked<%end if%>>KB</label>
<label><input type="radio" name="danwei" value="MB" <%if rs("danwei")="MB" then%>checked<%end if%>>MB</label>
<label><input type="radio" name="danwei" value="GB" <%if rs("danwei")="GB" then%>checked<%end if%>>GB</label></td>
      </tr>
    <tr bgcolor="#ffffff">
      <td height="12" align="right" class="td">下载地址</td>
      <td width="89%" class="td"><table width="48%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="39%"><input name="url" type="text" value="<%=rs("url")%>" size="50" /></td>
    <td width="61%"><iframe src="up.asp?formname=myform&inputname=url&uploadstyle=Download" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
  </tr>
</table>
</td>

    </tr>
      <tr>
        <td height="25" align="right" class="td">程序介绍</td>
        <td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"><%=replace(rs("body"),"'","&#39;")%></textarea>
</td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
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
title=request.form("title")
ssfl=request.form("ssfl")
yuyan=request.form("yuyan")
yxpt=request.form("yxpt")
tjdj=request.form("tjdj")
daxiao=request.form("daxiao")
danwei=request.form("danwei")
body=request.form("body")
url=request.form("url")
qx=request.form("qx")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from download where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
if title=""  then 
response.Write("<script language=javascript>alert('下载名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if yxpt=""  then 
response.Write("<script language=javascript>alert('请选择程序运行平台!');history.go(-1)</script>") 
response.end 
end if
if daxiao=""  then 
response.Write("<script language=javascript>alert('请填写程序大小!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("daxiao")) then
response.write("<script>alert(""程序大小必须为数字！""); history.go(-1);</script>")
response.end
end if
if url=""  then 
response.Write("<script language=javascript>alert('下载地址不能为空!');history.go(-1)</script>") 
response.end 
end if
if body=""  then 
response.Write("<script language=javascript>alert('内容不能为空!');history.go(-1)</script>") 
response.end 
end if
rs("title")=title
rs("ssfl")=ssfl
rs("yuyan")=yuyan
rs("yxpt")=yxpt
rs("daxiao")=daxiao
rs("tjdj")=tjdj
rs("danwei")=danwei
rs("body")=body
rs("url")=url
rs("qx")=qx
rs.update 
rs.close 
response.write "<script>alert('修改成功！');window.location.href='admin_download.asp';</script>" 
end if
%> 