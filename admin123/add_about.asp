<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(9)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>����ҳ��</title>		
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
<form  name="add" method="post" action="add_about.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">����ҳ��</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <tr>
        <td width="13%" height="28" align="right" class="td">��ҳ����</td>
        <td width="87%"  class="td"><input name="title" type="text" size="50"  /></td>
      </tr>
      <tr>
        <td width="13%" height="25" align="right" class="td">��ҳ�ؼ���</td>
        <td class="td"><input name="keywords" type="text" size="50"  /> | �ֿ�</td>
      </tr>
      <tr>
        <td width="13%" height="25" align="right" class="td">��ҳ����</td>
        <td class="td"><input name="description" type="text" size="50"  /></td>
      </tr>

        <tr>
        <td height="13" align="right" class="td">�Ƿ�ǰ̨��ʾ</td>
        <td class="td"><label>
          <input name="xianshi" type="radio"  value="1" checked="checked" />��ʾ
          <input type="radio" name="xianshi"  value="0" />����ʾ</label></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">ҳ������</td>
        <td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"></textarea></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
      </tr>
    </table></td>
  </tr>
</table></form>

</body>
</html>
<%
if request("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from about"
rs.open sql,conn,1,3
title=request.form("title")
body=request.form("body")
keywords=request.form("keywords")
description=request.form("description")
xianshi=request.form("xianshi")
px_id=request.form("px_id")
if title=""  then 
response.Write("<script language=javascript>alert('�������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('����ID����Ϊ��Ŷ!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
if body=""  then 
response.Write("<script language=javascript>alert('���ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
rs("title")=title
rs("body")=body
rs("keywords")=keywords
rs("description")=description
rs("xianshi")=xianshi
rs("px_id")=px_id
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('ҳ�����ӳɹ���');window.location.href='admin_about.asp';</script>" 
end if
%>