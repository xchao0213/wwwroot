<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<% 
call chkAdmin(12)
exec="select * from [case] where id="& request.QueryString("id") 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
ImagePath=rs("ImagePath")
img=rs("img") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>�޸İ���</title>
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

<form  name="add" method="post" onsubmit="return check(this)" action="xiugai_case.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸İ���</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <tr>
        <td width="12%" height="28" align="right" class="td"><font color="#FF0000">*</font>��������</td>
        <td width="88%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="50"  />
          <label><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></label></td>
      </tr>
      <tr>
        <td width="12%" height="13" align="right" class="td">��������</td>
        <td class="td"><select name="ssfl" id="select">
			<%
            dim rsc
            set rsc=server.CreateObject("adodb.recordset")
            rsc.open "select * from case_class",conn,1,1
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
            </select></td>
      </tr>
       <tr>
        <td width="12%" height="12" align="right" class="td">��������</td>
        <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="50" />  ����д������ַ����:http://www.fbiao.com </td>
      </tr>
<tr>
<td height="12" align="right" class="td">��������ͼ</td>
<td class="td"><table width="55%" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td width="32%"><input  name="img" type="text" id="img" size="50"value="<%=img%>" /></td>
  <td width="10%" align="center"><input name="button" type="button" class="btn2" id="flash" value="ѡȡ���ϴ�ͼƬ" /></td>
<td width="58%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=case" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
</tr>
</table></td>
</tr>
<tr>
<TD align=right height=30 class="td">�ϴ�ͼƬͼƬ</td>
<td colspan="3" class="td"><input name="ImagePath" type="hidden" id="ImagePath" value="<%=ImagePath%>" size="60">
<div id="pw">
<%
if ImagePath<>"" then 
dim ii,images
	images=split(ImagePath,"|")
	for ii=0 to ubound(images)
	response.write "<div id=""img"&ii&""" class=""imgDiv""><a href=""javascript:void(0)"" onclick=""setimg('"&images(ii)&"')""><img border=""0"" src="""&images(ii)&"""></a><br><label><input type=""radio"" value="""&images(ii)&""" onclick=""setimg('"&images(ii)&"')"" name=""imgradio"">��Ϊ����ͼ</label> <a href=""javascript:dropThisDiv('img"&ii&"','"&images(ii)&"')"">ɾ��</a></div>"
next
end if
%></div>
<script type="text/javascript">
//parent.parent.FCKeditorAPI.GetInstance('content')
// ���ñ༭��������
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
doChange(document.getElementById("ImagePath"),document.getElementById("ImageFileList"))
</script>
</td>
</tr>
<tr>
        <td height="13" align="right" class="td">����ʱ��</td>
        <td class="td"><input name="data" type="text" value="<%=rs("data")%>" size="30"  /><input name=button type=button class="btn2" onclick="document.add.data.value='<%=(now)%>'" value="��Ϊ��ǰʱ��"></td>
      </tr>
     <tr>
        <td height="25" align="right" class="td">�������� <font color="#FF0000">*</font></td>
        <td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"><%=replace(rs("body"),"'","&#39;")%></textarea></td>
      </tr>
     <tr>
        <td height="13" align="right" class="td">����ͼƬ</td>
        <td class="td"><label><input name="local" type="checkbox"  id="local" value="1">����Զ��ͼƬ������</label>
        <span id="one">(����Զ������ͼƬ��Ҫʱ�䣬���ܻ��ӳ�һ���!)</span></td>
      </tr>
      <tr>
        <td height="13" align="right" class="td">�Ƿ��Ƽ�</td>
        <td class="td"><label>
<input type="radio" name="tuijian" value="0" <%if rs("tuijian")=0 then%>checked<%end if%>>���Ƽ��� 
<input type="radio" name="tuijian" value="1" <%if rs("tuijian")=1 then%>checked<%end if%>>�Ƽ�
</label></td>
      </tr>
       <tr>
          <td height="12" align="right" class="td">&nbsp;</td>
          <td class="td"><input type="submit" name="button" id="button" value="ȷ���޸�"  class="btn"/></td>
        </tr>
    </table>
</td>
  </tr>
</table></form>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
url=request.form("url")
body=request.form("body")
color=request.form("color")
ssfl=request.form("ssfl")
img=request.form("img")
ImagePath=request.form("ImagePath")
data=request.form("data")
tuijian=request.form("tuijian")
local=request.Form("local")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from [case] where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if
if title=""  then 
response.Write("<script language=javascript>alert('�������ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if body=""  then 
response.Write("<script language=javascript>alert('�������ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if

if local=1 then
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	uploaddate=right(year(date),4)&right("00"&month(date),2)
	uploadpath="../uploadfile/image/case/"
	If Not fso.FolderExists(Server.mappath(uploadpath)) Then
	  fso.CreateFolder(Server.mappath(uploadpath))'����Ŀ¼
	 End If
	s_img=ReplaceRemoteUrl(img,uploadpath,sFileExt)
	s_body=ReplaceRemoteUrl(body,uploadpath,sFileExt)
else
	s_img=img
	s_body=body
end if
rs("title")=title
rs("url")=url
rs("body")=s_body
rs("color")=color
rs("ssfl")=ssfl
rs("img")=s_img
rs("ImagePath")=ImagePath
rs("data")=data
rs("tuijian")=tuijian
rs.update 
rs.close 
response.write "<script>alert('�����޸ĳɹ���');window.location.href='admin_case.asp';</script>" 
end if%>
</body>
</html>
