<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(20)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>增加下载</title>
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
<form  name="myform" method="post" action="add_download.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加下载</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <tr>
        <td width="10%" height="28" align="right" class="td">下载名称</td>
        <td colspan="2"  class="td">
          <input name="title" type="text" size="50"  /></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">所属分类</td>
        <td colspan="2" class="td"><select name="ssfl" id="select">
			  <% set rsc=server.CreateObject("adodb.recordset")
              rsc.open "select * from download_fl",conn,1,1
              while not rsc.eof
              response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
              rsc.movenext
              wend
              rsc.close
              set rsc=nothing%></select> 下载权限：<select name="qx" id="select2">
          <option value="1" checked>所有游客</option>
          <option value="9">本站会员</option>
        </select></td>
      </tr>
      <tr>
        <td width="10%" height="7" align="right" class="td">程序语言</td>
        <td colspan="2" class="td">
          <label><input name="yuyan" type="radio"value="简体中文" checked="checked" />简体中文</label>
          <label><input type="radio" name="yuyan"value="繁体中文 " />繁体中文 </label>
          <label><input type="radio" name="yuyan"value="English " />English </label>
          <label><input type="radio" name="yuyan"value="多国语言" />多国语言</label></td>
      </tr>
      <tr>
        <td width="10%" height="6" align="right" class="td">运行平台</td>
        <td colspan="2" class="td">
<label><input name="yxpt" type="checkbox" value="XP"/>XP </label>
<label><input name="yxpt" type="checkbox" value="Vista" />Vista</label>
<label><input name="yxpt" type="checkbox" value="Linux" />Linux</label>
<label><input name="yxpt" type="checkbox" value="Win7" />Win7</label>
<label><input name="yxpt" type="checkbox" value="Win8" />Win8</label>
<label><input name="yxpt" type="checkbox" value="android" />android</label>
<label><input name="yxpt" type="checkbox" value="IOS" />IOS</label></td>
      </tr>
      <tr>
        <td width="10%" height="12" align="right" class="td">推荐等级</td>
        <td colspan="2" class="td"><label>
          <select name="tjdj" id="select2">
            <option value="★☆☆☆☆">★☆☆☆☆</option>
            <option value="★★☆☆☆">★★☆☆☆</option>
            <option value="★★★☆☆">★★★☆☆</option>
            <option value="★★★★☆">★★★★☆</option>
            <option value="★★★★★">★★★★★</option>
          </select>
        </label></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">程序大小</td>
        <td colspan="2" class="td"><input name="daxiao" type="text" size="10"  />
          
<label><input name="danwei" type="radio" value="KB" checked="checked" />KB</label>
<label><input type="radio" name="danwei" value="MB" />MB</label>
<label><input type="radio" name="danwei" value="GB" />GB</label></td>
      </tr>
    <tr>
      <td height="25" align="right" class="td">下载地址</td>
      <td width="23%" class="td"><input type=text name=url size=40></td>
      <td width="67%" class="td"><iframe src="up.asp?formname=myform&inputname=url&uploadstyle=Download" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">程序介绍</td>
        <td colspan="2" class="td"><textarea name="body" style="width:700px;height:200px;visibility:hidden;"></textarea></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td colspan="2" class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
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
sql="select * from download"
rs.open sql,conn,1,3
title=request.form("title")
ssfl=request.form("ssfl")
yuyan=request.form("yuyan")
yxpt=request.form("yxpt")
daxiao=request.form("daxiao")
tjdj=request.form("tjdj")
danwei=request.form("danwei")
body=request.form("body")
url=request.form("url")
qx=request.form("qx")
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
rs.addnew
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
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('下载增加成功！');window.location.href='add_download.asp';</script>" 
end if
%>