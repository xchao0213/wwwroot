<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(14)
exec="select * from Products where id="& request.QueryString("id") 
set rsa=server.createobject("adodb.recordset") 
rsa.open exec,conn,1,1
ImagePath=rsa("ImagePath")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>�޸Ĳ�Ʒ</title>
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
<script language = "JavaScript">
<%
set rs = server.createobject("adodb.recordset")
rs.open "select * from smallclass order by smallclassid asc",conn,1,1
%>
var onecount;
onecount=0;
subcat = new Array();
<%
   i = 0
   Do While Not Rs.eof 
%>
subcat[<%=i%>] = new Array("<%= Trim(Rs("SmallClassName"))%>","<%= Rs("BigClassID")%>","<%= Rs("SmallClassID")%>");
<%
        i = i + 1
        Rs.MoveNext
        Loop
        Rs.Close
%>
  
onecount=<%=i%>;

function changelocation(locationid,formname)
    {
    formname.SmallClassID.length = 0; 

    var locationid = locationid;
    var i;
    for (i = 0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            {
             formname.SmallClassID.options[formname.SmallClassID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script>
<style type="text/css">
<!--
#pw{height: auto; overflow:hidden;}
.imgDiv{width:117px; height:105px;float:left; margin:2px;border:1px #CCC solid ;padding:1px}
.imgDiv img{ width:115px; height:80px; margin:0 auto;border:1px #CCC solid}
.imgDiv img:hover{ border:1px #CC0000 solid}
-->
</style>
</head>
<body>
<form  name="form1" method="post" action="xiugai_products.asp?id=<%=rsa("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸Ĳ�Ʒ</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="13%" height="28" align="right" class="td"><font color="#FF0000">*</font>��Ʒ����</td>
        <td width="87%"  class="td">
          <input name="title" type="text" value="<%=rsa("title")%>" size="50"  />
          <label><input name="id" type="hidden" id="id" value="<%=rsa("id")%>" /></label> 
		  </td>
        </tr>
        
        <tr>
<td height="12" align="right" class="td">��������ͼ</td>
<td class="td"><table width="55%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="31%"><input type=text id="img" name=img size=50  value="<%=rsa("img")%>"></td>
<td width="9%"><input type="button" class="btn2" id="flash" value="ѡȡ���ϴ�ͼƬ" /></td>
<td width="60%"><iframe src="up.asp?formname=form1&inputname=img&uploadstyle=Products" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
</tr>
</table>
</td>
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
</script></td>
</tr>
<tr>
    <td width="13%" height="13" align="right" class="td">��Ʒ����</td>
     <td class="td"><select name="BigClassID" id="BigClassID" onChange="changelocation(document.form1.BigClassID.options[document.form1.BigClassID.selectedIndex].value,document.form1)">
    <option value="">��ѡ�����</option>
    <%
    i = 0
    Set RsBig = Conn.Execute("SELECT * FROM BigClass ORDER BY px_id")
    Do While Not RsBig.Eof
        If i = 0 Then BigClassID = RsBig(0)
            IF Cstr(rsa("BigClassID"))=Cstr(RsBig(0)) THEN SelectStr = " selected" ELSE SelectStr = ""
            Response.Write("<option value='"&RsBig(0)&"'"& SelectStr &">"&RsBig(1)&"</option>")
        i = i + 1
        RsBig.MoveNext
    Loop
    RsBig.Close:Set RsBig = Nothing
    %>
    </select>
    <select name="SmallClassID" id="SmallClassID">
    <option value="">��ѡ��С��</option>
    <%
    IF rsa("BigClassID")<>"" THEN
        Set RsSmall = Conn.Execute("SELECT * FROM SmallClass WHERE BigClassID="&rsa("BigClassID"))
        Do While Not RsSmall.Eof
            IF Cstr(rsa("SmallClassID"))=Cstr(RsSmall(0)) THEN SelectStr = " selected" ELSE SelectStr = ""
            Response.Write("<option value='"&RsSmall(0)&"'"& SelectStr &">"&RsSmall(1)&"</option>")
            RsSmall.MoveNext
            Loop
        RsSmall.Close:Set RsSmall = Nothing
    End IF
    %>
    </select>        </td>
  </tr>      
		<tr>
         <td height="25" align="right" class="td">�ؼ���</td>
         <td class="td"><input name="key" type="text" value="<%=rsa("key")%>"  size="50" />����ؼ�������,���� </td>
       </tr>
		<tr>
         <td height="25" align="right" class="td">�Ա���ַ</td>
         <td class="td"><input name="taourl" type="text" value="<%=rsa("taourl")%>"  size="50" />�Ա�����ı�����ַ</td>
       </tr>
       <tr>
        <td height="13" align="right" class="td">��������</td>
        <td class="td"><% 
exec="select * from admin where id="&session("admin")&" "
set rsn=server.createobject("adodb.recordset") 
rsn.open exec,conn,1,1 
%> <input name="zz" type="text" value="<%=rsa("zz")%>" size="50"  /></td>
      </tr>
<tr>
<td height="25" align="right" class="td">��Ʒ�ͺ�</td>
<td class="td"><label><input name="cpxh" type="text" value="<%=rsa("cpxh")%>"  size="50" /></label></td>
</tr>
<tr>
<td height="25" align="right" class="td">��Ʒ���</td>
<td class="td"><label><input name="cpgg" type="text" value="<%=rsa("cpgg")%>"  size="50" /></label></td>
</tr>
<tr>
      <tr>
         <td height="25" align="right" class="td">��Ʒ�۸�</td>
         <td class="td"><input name="jiage" type="text" value="<%=rsa("jiage")%>"  size="20" /> ���Żݼ�<input name="zkj" type="text" value="<%=rsa("zkj")%>"  size="20" /> ��</td>
       </tr>
       <tr>
         <td height="25" align="right" class="td">��ݷ���</td>
         <td class="td"><input name="kdf" type="text" value="<%=rsa("kdf")%>"  size="20" /> ����Ʊ��<input name="piaodian" type="text" value="<%=rsa("piaodian")%>"  size="20" /> ��</td>
       </tr>
	   <tr>
         <td height="25" align="right" class="td">�Ƽ��Ǽ�</td>
         <td class="td"><select name="xin">
          <option value="������" <%if rsa("xin")="������" then%>selected<%end if%>>������</option>
          <option value="������" <%if rsa("xin")="������" then%>selected<%end if%>>������</option>
		  <option value="������" <%if rsa("xin")="������" then%>selected<%end if%>>������</option>
		  <option value="������" <%if rsa("xin")="������" then%>selected<%end if%>>������</option>
		  <option value="������" <%if rsa("xin")="������" then%>selected<%end if%>>������</option>
        </select>
          </td>
       </tr>
	  <tr>
        <td height="13" align="right" class="td">����ʱ��</td>
        <td colspan="2" class="td"><input name="data" type="text" value="<%=rsa("data")%>" size="30"  /> <input name=button type=button class="btn2" onclick="document.form1.data.value='<%=(now)%>'" value="��Ϊ��ǰʱ��"></td>
	</tr> 
     <tr>
        <td height="25" align="right" class="td">��Ʒ���� <font color="#FF0000">*</font></td>
        <td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"><%=replace(rsa("body"),"'","&#39;")%></textarea></td>
      </tr>
       <tr>
        <td height="13" align="right" class="td">�Ƿ��Ƽ�</td>
        <td class="td"><label>
<input type="radio" name="tuijian" value="0" <%if rsa("tuijian")=0 then%>checked<%end if%>>���Ƽ��� 
<input type="radio" name="tuijian" value="1" <%if rsa("tuijian")=1 then%>checked<%end if%>>�Ƽ�</label></td>
      </tr>
     <tr>
         <td height="12" align="right" class="td">&nbsp;</td>
         <td class="td"><input type="submit" name="button" id="button" value="ȷ���޸�"  class="btn"/></td>
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
BigClassID=request.form("BigClassID")
SmallClassID=request.form("SmallClassID")
key=request.form("key")
zz=request.form("zz")
zkj=request.form("zkj")
img=request.form("img")
body=request.form("body")
tuijian=request.form("tuijian")
jiage=request.form("jiage")
xin=request.form("xin")
data=request.form("data")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from products where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if
if title=""  then 
response.Write("<script language=javascript>alert('��Ʒ���ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if BigClassID="" then
response.Write("<script language=javascript>alert('��ѡ���Ʒ����!');history.go(-1)</script>")
response.End()
end if
if SmallClassID="" then
SmallClassID=0
end if
if body=""  then 
response.Write("<script language=javascript>alert('��Ʒ���ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
rs("title")=title
rs("BigClassID")=BigClassID
rs("SmallClassID")=SmallClassID
rs("img")=img
rs("key")=key
rs("zz")=zz
rs("zkj")=zkj
rs("body")=body
rs("tuijian")=tuijian
rs("taourl")=request.form("taourl")
rs("ImagePath")=request.form("ImagePath")
rs("cpxh")=request.form("cpxh")
rs("cpgg")=request.form("cpgg")
rs("kdf")=request.form("kdf")
rs("piaodian")=request.form("piaodian")
rs("jiage")=jiage
rs("xin")=xin
rs("data")=data
rs.update 
rs.close 
response.write "<script>alert('��Ʒ�޸ĳɹ���');window.location.href='admin_products.asp';</script>" 
end if
%> 