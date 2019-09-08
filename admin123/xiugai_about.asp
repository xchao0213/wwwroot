<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(9)%>
<%id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误00！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from about where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>修改页面</title>
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
</head>
<body>
<form  name="add" method="post" action="xiugai_about.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#666666">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改页面</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="12%" height="28" align="right" class="td">网页标题</td>
        <td width="88%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="30"  />
          <input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr>
        <td width="12%" height="25" align="right" class="td">网页关键词</td>
        <td class="td"><input name="keywords" type="text" value="<%=rs("keywords")%>" size="50"  /> 
        | 分开</td>
      </tr>
      <tr>
        <td width="12%" height="25" align="right" class="td">网页描述</td>
        <td class="td"><input name="description" type="text" value="<%=rs("description")%>" size="50"  /> </td>
      </tr>

        <tr>
        <td height="13" align="right" class="td">是否前台显示</td>
        <td class="td"><label>
<input type="radio" name="xianshi" value="0" <%if rs("xianshi")=0 then%>checked<%end if%>>不显示　 
<input type="radio" name="xianshi" value="1" <%if rs("xianshi")=1 then%>checked<%end if%>>显示
</label></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="30"  /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">页面内容</td>
        <td class="td">
        <textarea name="body" style="width:700px;height:400px;visibility:hidden;"><%=replace(rs("body"),"'","&#39;")%></textarea>
      </td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
      </tr>
    </table></td>
  </tr>
</table></form>
</body>
</html>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
body=request.form("body")
keywords=request.form("keywords")
description=request.form("description")
xianshi=request.form("xianshi")
px_id=request.form("px_id")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from about where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
	if title=""  then 
response.Write("<script language=javascript>alert('标题名称不能为空!');history.go(-1)</script>") 
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
if body=""  then 
response.Write("<script language=javascript>alert('内容不能为空!');history.go(-1)</script>") 
response.end 
end if
rs("title")=title
rs("body")=body
rs("keywords")=keywords
rs("description")=description
rs("xianshi")=xianshi
rs("px_id")=px_id
rs.update 
rs.close 
response.write "<script>alert('修改成功！');window.location.href='admin_about.asp';</script>" 
end if
%> 
