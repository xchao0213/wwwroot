<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(22)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>���Ӱ���</title>
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

/*����ͼƬ����ͼ*/
function setimg(t)
{
document.getElementById("img").value=t
}
/*ɾ��ͼƬ*/
function dropThisDiv(t,p)
{
document.getElementById(t).style.display='none'
var str =document.getElementById("ImagePath").value;
var arr = str.split("|");
var nstr="";
for (var i=0; i<arr.length; i++)
{
	if(arr[i]!=p)
	{
		if (nstr!="")
		{
			nstr=nstr+"|";
		}		
		nstr=nstr+arr[i]
	}
}
document.getElementById("ImagePath").value=nstr;
}
/*��������ͼ��*/
function valMsg(strValue){
	if (strValue=="") return;
	var objLinkUpload =document.getElementsByName("ImagePath")[0];
	if (objLinkUpload){
		if (objLinkUpload.value!=""){
			objLinkUpload.value = objLinkUpload.value + "|";
		}
		objLinkUpload.value = objLinkUpload.value + strValue;
	}
}
/*���ͼƬ����*/
KindEditor.ready(function(K) {
var editor = K.editor({
	fileManagerJson : '../zycheditor/asp/file_manager_json.asp'
});
K('#flash').click(function() {
	editor.loadPlugin('filemanager', function() {
		editor.plugin.filemanagerDialog({
			viewType : 'VIEW',
			dirName : 'image',
			urlType : 'relative',
			clickFn : function(url, title) {
				K('#img').val(url);
				var pw=document.getElementById("pw").innerHTML;
				var ids=url
				   ids=ids.substr(ids.lastIndexOf("/")+1,ids.lastIndexOf("."))
				   ids=ids.replace(".","")
				var div=pw+'<div class="imgDiv" id="'+ids+'"><label><img src="'+ url + '" border="0" /><br><input type="radio" name="imgradio" onclick="setimg(\''+url+'\')" value="'+url+'" />��Ϊ����ͼ</label> <a href="javascript:dropThisDiv(\''+ids+'\',\''+url+'\')">ɾ��</a></div>'
				
				K('#pw').html(div);
				K('#ImagePath').val(valMsg(url));
				editor.hideDialog();
			}
		});
	});
});  
  });
</script>
</head>
<body>
<form  name="add" method="post" action="add_photo.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="11%" height="28" align="right" class="td"><font color="#FF0000">*</font>������</td>
        <td width="89%"  class="td"><input name="title" type="text" size="40"  /></td>
      </tr>
      <tr>
        <td width="11%" height="13" align="right" class="td"><font color="#FF0000">*</font>������</td>
        <td class="td"><select name="ssfl" id="select">
					<%set rsc=server.CreateObject("adodb.recordset")
                    rsc.open "select * from photo_class",conn,1,1
                    while not rsc.eof
                    response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
                    rsc.movenext
                    wend
                    rsc.close
                    set rsc=nothing%>
                    </select></td>
      </tr>
          <tr>
        <td width="11%" height="12" align="right" class="td">��������</td>
        <td class="td"><input name="url" type="text" value="#" size="40" /> 
          ����д������ַ����:QQ�ռ�http://495573637.qzone.qq.com/ </td>
      </tr>
    <tr>
    <TD  height=30 align="right" valign="middle" class="td">����ͼ</td>
      <td colspan="3" class="td"><input name="img" type="text" id="img" size="40" value="" />
        <input name="button" type="button" class="btn2" id="flash" value="ѡȡ���ϴ�ͼƬ" /></td>
    </tr>
	<tr>
        <td width="11%" height="12" align="right" class="td"><font color="#FF0000">*</font>ͼƬ��</td>
        <td class="td"><input name="ImagePath" type="hidden" id="ImagePath" value="" size="60">
        <div id="result"></div> 
<div id="pw"></div>
<script type="text/javascript">
function SetEditorContents(EditorName, ContentStr) {
     var oEditor = FCKeditorAPI.GetInstance(EditorName) ;
	 oEditor.Focus();
     oEditor.InsertHtml("<img src="+"'"+ContentStr+"'"+"/>");	
}
function SetEditorPage(EditorName, ContentStr) {
     var oEditor = FCKeditorAPI.GetInstance(EditorName) ;
	 oEditor.Focus();
     oEditor.InsertHtml(ContentStr);	
}
</script>

<script type="text/javascript">
function dropThisDiv(t,p)
{
document.getElementById(t).style.display='none'
var str =document.getElementById("ImagePath").value;
var arr = str.split("|");
var nstr="";
for (var i=0; i<arr.length; i++)
{
	if(arr[i]!=p)
	{
		if (nstr!="")
		{
			nstr=nstr+"|";
		}		
		nstr=nstr+arr[i]
	}
}
document.getElementById("ImagePath").value=nstr;

doChange(document.getElementById("ImagePath"),document.getElementById("ImageFileList"))
}

function setimg(t)
{
document.getElementById("img").value=t

doChange(document.getElementById("ImagePath"),document.getElementById("ImageFileList"))
}

function doChange(objText, objDrop){
if (!objDrop) return;
var str = objText.value;
var arr = str.split("|");
var nIndex = objDrop.selectedIndex;
objDrop.length=1;
for (var i=0; i<arr.length; i++){
objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
}
objDrop.selectedIndex = nIndex;
}
</script>
</td>
      </tr>
    <tr>
    <TD height="30" align="right" class="td">�ϴ�ͼƬ</td>
      <td colspan="3" class="td"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=photo" frameborder="0" scrolling="no" width="450" height="25"></iframe></td>
    </tr>
     <tr>
        <td height="25" align="right" class="td">��Ƭ����</td>
        <td class="td"><textarea name="body" style="width:700px;height:200px;visibility:hidden;"></textarea></td>
      </tr>
       <tr>
        <td height="13" align="right" class="td">�Ƿ��Ƽ�</td>
        <td class="td"><label>
          <input name="tuijian" type="radio"  value="0" checked="checked" />���Ƽ�
          <input type="radio" name="tuijian"  value="1" />�Ƽ� </label></td>
      </tr>
     <tr>
         <td height="47" align="right" class="td">&nbsp;</td>
         <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
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
sql="select * from photo"
rs.open sql,conn,1,3
title=request.form("title")
ssfl=request.form("ssfl")
url=request.form("url")
img=request.form("img")
ImagePath=request.form("ImagePath")
body=request.form("body")
tuijian=request.form("tuijian")
if title=""  then 
response.Write("<script language=javascript>alert('���ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if ImagePath=""  then 
response.Write("<script language=javascript>alert('ͼƬ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if ssfl=""  then 
response.Write("<script language=javascript>alert('�����಻��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
rs("title")=title
rs("ssfl")=ssfl
rs("url")=url
rs("img")=img
rs("ImagePath")=ImagePath
rs("body")=body
rs("tuijian")=tuijian
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('���ӳɹ�������������ӣ�');window.location.href='admin_photo.asp';</script>"
end if
%>