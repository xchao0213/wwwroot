<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(10)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>��������</title>
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
	if(form.local.checked) {//�ж�ѡ��ʱִ��
		 // alert("��ַ����Ϊ��!");
		  document.getElementById('one').innerHTML='<font color="#FF0000">Զ��ͼƬ���ڱ��ػ�...��ȴ���</font>';
		  form.local.focus();
		  //return false;  //ע�ͺ����ִ��
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
<form  name="add" method="post" onsubmit="return check(this)" action="add_news.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <tr>
        <td width="14%" height="28" align="right" class="td"><font color="#FF0000">*</font>���ű���</td>
        <td width="86%"  class="td"><input name="title" type="text" size="50"  /></td></tr>
      <tr>
      <td height="12" align="right" class="td">��������ͼ</td>
      <td class="td"><table width="52%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%"><input type=text id="img" name=img size=50></td>
    <td width="12%"><input name="button2" type="button" id="flash" class="btn2" value="ѡȡ���ϴ�ͼƬ" /></td>
    <td width="58%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=News" frameborder="0" scrolling="no" width="300" height="25"></iframe></td>
  </tr>
</table></td>
    </tr>
  <tr>
        <td width="14%" height="25" align="right" class="td">�ؼ���</td>
        <td class="td"><input name="key" type="text"  size="50"  />�������,����</td>
      </tr>


      <tr>
        <td width="14%" height="13" align="right" class="td">���ŷ���</td>
        <td class="td"><select name="ssfl" id="select">
            <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from news_fl",conn,1,1
		  while not rsc.eof
			response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing%> </select> ���Ȩ�ޣ�<select name="qx" id="select2">
          <option value="1" checked>�����ο�</option>
          <option value="9">��վ��Ա</option>
        </select></td>
      </tr>
         <tr>
        <td width="14%" height="12" align="right" class="td">�ⲿ����</td>
        <td class="td"><input name="url" type="text" size="50"  /> 
          ����д����ֱ����ת���ⲿ���ӵ�ַ��������Ϣ����Ч��</td>
      </tr>
  <tr>
        <td width="14%" height="25" align="right" class="td">������Դ</td>
        <td class="td"><input name="ly" type="text" value="��վ" size="30"  /></td>
      </tr>
       <tr>
        <td height="13" align="right" class="td">��������</td>
        <td class="td"><% 
exec="select * from admin where id="&session("admin")&""
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
%> <input name="zz" type="text" value="<%=rs("zsname")%>" size="30"  /></td>
      </tr>
       <tr>
        <td height="13" align="right" class="td">����ʱ��</td>
        <td class="td"><input name="data" type="text" value="<%=(now)%>" size="30"  /></td>
      </tr>

     <tr>
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>��������</td>
        <td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"></textarea></td>
      </tr>
      <tr>
        <td height="13" align="right" class="td">����ͼƬ</td>
        <td class="td"><label><input name="local" type="checkbox"  id="local" value="1">����Զ��ͼƬ������</label> <span id="one">(����Զ������ͼƬ��Ҫʱ�䣬���ܻ��ӳ�һ���!)</span></td>
      </tr>
       <tr>
        <td height="13" align="right" class="td">�Ƿ��Ƽ�</td>
        <td class="td"><label>
          <input name="tuijian" type="radio"  value="0" checked="checked" />���Ƽ�
          <input type="radio" name="tuijian"  value="1" />�Ƽ� </label></td>
      </tr>
     <tr>
         <td height="12" class="td">&nbsp;</td>
         <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn" onclick="window.location.href='../admin/asptohtml/show.asp?flag=ShowNews&amp;id=<%=rs("id")%>'"/></td>
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
sql="select * from news"
rs.open sql,conn,1,3
title=request.form("title")
url=request.form("url")
body=request.form("body")
img=request.form("img")
ly=request.form("ly")
zz=request.form("zz")
color=request.form("color")
ssfl=request.form("ssfl")
tuijian=request.form("tuijian")
qx=request.form("qx")
local=request.Form("local")
if title=""  then 
response.Write("<script language=javascript>alert('���ű��ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if body=""  then 
response.Write("<script language=javascript>alert('�������ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
if local=1 then
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	uploaddate=right(year(date),4)&right("00"&month(date),2)
	uploadpath="../uploadfile/image/News/"
	s_img=ReplaceRemoteUrl(img,uploadpath,sFileExt)
	s_body=ReplaceRemoteUrl(body,uploadpath,sFileExt)
else
	s_img=img
	s_body=body
end if
rs("title")=title
rs("url")=url
rs("body")=s_body
rs("img")=s_img
rs("ly")=ly
rs("zz")=zz
rs("color")=color
rs("ssfl")=ssfl
rs("tuijian")=tuijian
rs("qx")=qx
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('�������ӳɹ�������������ӣ�');window.location.href='add_news.asp';</script>" 
end if
%>