<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(10)
exec="select * from culture where id="& request.QueryString("id") 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>修改文化</title>
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
<script type="text/javascript">
function check(form) {
	if(form.local.checked) {//判断选中时执行
		 // alert("地址不能为空!");
		  document.getElementById('one').innerHTML='<font color="#FF0000">远程图片正在本地化...请等待！</font>';
		  form.local.focus();
		  //return false;  //注释后继续执行
	 }
return true;
}

KindEditor.ready(function(K) {
  var editor = K.editor({
	  fileManagerJson : '../zycheditor/asp/file_manager_json.asp'
  });
  K('#flash').click(function() {
	  editor.loadPlugin('filemanager', function() {
		  editor.plugin.filemanagerDialog({
			  viewType : 'VIEW',
			  dirName : 'image',
			  clickFn : function(url, title) {
				  K('#img').val(url);
				  editor.hideDialog();
			  }
		  });
	  });
  });
});
</script>
</head>
<body>
<form  name="add" method="post" onsubmit="return check(this)" action="xiugai_culture.asp?id=<%=rs("id")%>&modification=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改文化</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <tr>
        <td width="13%" height="28" align="right" class="td"><font color="#FF0000">*</font>文化标题</td>
        <td width="87%"  class="td"><input name="title" type="text" value="<%=rs("title")%>" size="50"  /></td>
      </tr>
 <tr>
      <td height="12" align="right" class="td">标题略缩图</td>
      <td class="td"><table width="51%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="32%"><input type=text id="img" name=img size=50  value="<%=rs("img")%>"></td>
    <td width="9%"><input name="button2" type="button" id="flash" class="btn2" value="选取已上传图片" /></td>
    <td width="59%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=Culture" frameborder="0" scrolling="no" width="300" height="25"></iframe></td>
  </tr>
</table>      </td>
    </tr>
    
     <tr>
        <td width="13%" height="25" align="right" class="td">关键字</td>
        <td class="td"><input name="key" type="text" value="<%=rs("key")%>" size="50"  /></td>
      </tr>
      <tr>
        <td width="13%" height="13" align="right" class="td">文化分类</td>
        <td class="td"><select name="ssfl" id="select">
          <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from culture_fl",conn,1,1
		  while not rsc.eof
		    if rs("ssfl")=rsc("id") then
			response.Write("<option value=""" & rsc("id") & """ selected>" & rsc("title") & "</option>")
			else
			response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing%> 
		</select>
		 浏览权限：<select name="qx" id="select2">
          <option value="1" <%if rs("qx")="1" then%>selected<%end if%>>所有游客</option>
          <option value="9" <%if rs("qx")="9" then%>selected<%end if%>>本站会员</option>
         </select></td>
      </tr>
       <tr>
        <td width="13%" height="12" align="right" class="td">外部链接</td>
        <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="50"  /> 如填写，则直接跳转到外部链接地址。</td>
      </tr>
     <tr>
        <td width="13%" height="25" align="right" class="td">文化来源</td>
        <td class="td"><input name="ly" type="text" value="<%=rs("ly")%>" size="30"  /></td>
      </tr>
   <tr>
        <td height="13" align="right" class="td">发布作者</td>
        <td class="td"> <input name="zz" type="text" value="<%=rs("zz")%>" size="30"  /></td>
      </tr>

<tr>
        <td height="13" align="right" class="td">发布时间</td>
        <td class="td"><input name="data" type="text" value="<%=rs("data")%>" size="30"  /><input name=button type=button class="btn2" onclick="document.add.data.value='<%=(now)%>'" value="设为当前时间"></td>
      </tr>

     <tr>
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>文化内容</td>
        <td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"><%=replace(rs("body"),"'","&#39;")%></textarea></td>
      </tr>
      <tr>
        <td height="13" align="right" class="td">外链图片</td>
        <td class="td"><label><input name="local" type="checkbox"  id="local" value="1">保存远程图片到本地</label> <span id="one">(由于远程下载图片需要时间，可能会延迟一会儿!)</span></td>
      </tr>
        <tr>
        <td height="13" align="right" class="td">是否推荐</td>
        <td class="td"><label>
<input type="radio" name="tuijian" value="0" <%if rs("tuijian")=0 then%>checked<%end if%>>不推荐　 
<input type="radio" name="tuijian" value="1" <%if rs("tuijian")=1 then%>checked<%end if%>>推荐
</label></td>
      </tr>
       <tr>
        <td height="12" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="确认修改"  class="btn"/></td>
        </tr>
    </table>
</td>
  </tr>
</table></form>
</body>
</html>
<% 
if Request.QueryString("modification")="ok" then
id=request("id")
title=request.form("title")
key=request.form("key")
fb=request.form("fb")
url=request.form("url")
body=request.form("body")
color=request.form("color")
ssfl=request.form("ssfl")
fl=request.form("fl")
ly=request.form("ly")
zz=request.form("zz")
data=request.form("data")
img=request.form("img")
qx=request.form("qx")
tuijian=request.form("tuijian")
local=request.Form("local")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from culture where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if
if title=""  then 
response.Write("<script language=javascript>alert('文化标题不能为空!');history.go(-1)</script>") 
response.end 
end if
if ssfl=""  then 
response.Write("<script language=javascript>alert('文化分类不能为空!');history.go(-1)</script>") 
response.end 
end if
if body=""  then 
response.Write("<script language=javascript>alert('文化内容不能为空!');history.go(-1)</script>") 
response.end 
end if
if local=1 then
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	uploaddate=right(year(date),4)&right("00"&month(date),2)
	uploadpath="../uploadfile/image/Culture/"
	s_img=ReplaceRemoteUrl(img,uploadpath,sFileExt)
	s_body=ReplaceRemoteUrl(body,uploadpath,sFileExt)
else
	s_img=img
	s_body=body
end if
rs("title")=title
rs("key")=key
rs("url")=url
rs("body")=s_body
rs("color")=color
rs("ssfl")=ssfl
rs("ly")=ly
rs("zz")=zz
rs("data")=data
rs("img")=s_img
rs("tuijian")=tuijian
rs("qx")=qx
rs.update 
rs.close 
response.write "<script>alert('文化修改成功！');window.location.href='admin_culture.asp';</script>" 
end if
%> 
