<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(14)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>增加产品</title>
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
                          K('#img').val(url);
                          editor.hideDialog();
                      }
                  });
              });
          });
      });
  </script>
<script language = "JavaScript">
<%
Set Rs = Server.CreateObject("Adodb.Recordset")
Rs.Open "SELECT * FROM SmallClass ORDER BY SmallClassID asc",conn,1,1
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
<form  name="form1" method="post" action="add_products.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加产品</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
<tr>
<td width="15%" height="28" align="right" class="td"><font color="#FF0000">*</font>产品标题</td>
<td width="85%"  class="td">
<input name="title" type="text" size="50"  />
</td>
</tr>
<tr>
<td height="12" align="right" class="td">略缩图</td>
<td class="td"><table width="58%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="21%"><input type=text id="img" name=img size=50></td>
<td width="21%"><input type="button" class="btn2" id="filemanager" value="选取已上传图片" /></td>
<td width="58%"><iframe src="up.asp?formname=form1&inputname=img&uploadstyle=Products" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
</tr>
</table>
</td>
</tr>
<tr>
<TD align=right height=30 class="td">上传图片图片</td>
<td colspan="3" class="td"><input name="ImagePath" type="hidden" id="ImagePath" value="<%=ImagePath%>" size="50">
<div id="pw"></div>
<script type="text/javascript">
//parent.parent.FCKeditorAPI.GetInstance('content')
// 设置编辑器中内容
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
<div></div></td>
</tr>
<tr>
<td width="15%" height="13" align="right" class="td">产品分类</td>
<td class="td"><select name="BigClassID" id="BigClassID" onChange="changelocation(document.form1.BigClassID.options[document.form1.BigClassID.selectedIndex].value,document.form1)">
<option value="">请选择大类</option>
<%
i = 0
Set RsBig = Conn.Execute("SELECT * FROM BigClass ORDER BY px_id")
Do While Not RsBig.Eof
If i = 0 Then BigClassID = RsBig(0)
Response.Write("<option value="""&RsBig(0)&""">"&RsBig(1)&"</option>")
i = i + 1
RsBig.MoveNext
Loop
Set RsBig = Nothing%>
</select>
<select name="SmallClassID" id="SmallClassID">
<option value="">请选择小类</option>
<%If BigClassID <> "" Then
Set RsSmall = Conn.Execute("SELECT * FROM SmallClass WHERE BigClassID = "&BigClassID)
Do While Not RsSmall.Eof
Response.Write("<option value="""&RsSmall(0)&""">"&RsSmall(1)&"</option>")
RsSmall.MoveNext
Loop
End if
Set RsSmall = Nothing%>
</select></td>
</tr>
<tr>
<td height="25" align="right" class="td">关键字</td>
<td class="td"><input name="key" type="text"  size="50" />多个关键字请用,隔开</td>
</tr>
  <tr>
   <td height="25" align="right" class="td">淘宝地址</td>
   <td class="td"><input name="taourl" type="text" value="<%=rsa("taourl")%>"  size="50" />淘宝网店的宝贝地址</td>
 </tr>
<tr>
<td height="25" align="right" class="td">产品型号</td>
<td class="td"><label><input name="cpxh" type="text"  size="50" /></label></td>
</tr>
<tr>
<td height="25" align="right" class="td">产品规格</td>
<td class="td"><label><input name="cpgg" type="text"  size="50" /></label></td>
</tr>
<tr>
<td height="25" align="right" class="td">产品价格</td>
<td class="td"><label><input name="jiage" type="text" value="0"  size="20" /> ￥</label>优惠价
<label><input name="zkj" type="text" value="0"  size="20" /> ￥</label></td>
</tr>
<tr>
 <td height="25" align="right" class="td">快递费用</td>
 <td class="td"><input name="kdf" type="text" value=""  size="20" /> ￥发票点<input name="piaodian" type="text" value=""  size="20" /> ￥</td>
</tr>
<tr>
<td height="25" align="right" class="td">推荐星级</td>
<td class="td"><select name="xin">
<option value="★☆☆☆☆">★☆☆☆☆</option>
<option value="★★☆☆☆">★★☆☆☆</option>
<option value="★★★☆☆" selected="selected">★★★☆☆</option>
<option value="★★★★☆">★★★★☆</option>
<option value="★★★★★">★★★★★</option>
</select>
</td>
</tr> 
<tr>
<td height="13" align="right" class="td">发布作者</td>
<td class="td"><% 
exec="select * from admin where id="&session("admin")&" "
set rsn=server.createobject("adodb.recordset") 
rsn.open exec,conn,1,1 %><input name="zz" type="text" value="admin" size="20"/></td>
</tr>
<tr>
<td height="25" align="right" class="td">产品介绍 <font color="#FF0000">*</font></td>
<td class="td"><textarea name="body" style="width:700px;height:400px;visibility:hidden;"></textarea></td>
</tr>
<tr>
<td height="13" align="right" class="td">是否推荐</td>
<td class="td"><label>
<input name="tuijian" type="radio"  value="0" checked="checked" />不推荐
<input type="radio" name="tuijian"  value="1" />推荐 </label></td>
</tr>
<tr>
<td height="12" align="right" class="td">&nbsp;</td>
<td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
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
sql="select * from Products"
rs.open sql,conn,1,3
title=request.form("title")
BigClassID=request.form("BigClassID")
SmallClassID=request.form("SmallClassID")
img=request.form("img")
zz=request.form("zz")
body=request.form("body")
tuijian=request.form("tuijian")
jiage=request.form("jiage")
zkj=request.form("zkj")
xin=request.form("xin")
if title=""  then 
response.Write("<script language=javascript>alert('产品标题不能为空!');history.go(-1)</script>") 
response.end 
end if
if BigClassID="" then
response.Write("<script language=javascript>alert('请选择产品分类!');history.go(-1)</script>")
response.End()
end if
if SmallClassID="" then
SmallClassID=0
end if
if body=""  then 
response.Write("<script language=javascript>alert('产品内容不能为空!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
rs("title")=title
rs("BigClassID")=BigClassID
rs("SmallClassID")=SmallClassID
rs("img")=img
rs("zz")=zz
rs("body")=body
rs("tuijian")=tuijian
rs("taourl")=request.form("taourl")
rs("ImagePath")=request.form("ImagePath")
rs("cpxh")=request.form("cpxh")
rs("cpgg")=request.form("cpgg")
rs("kdf")=request.form("kdf")
rs("piaodian")=request.form("piaodian")
rs("zkj")=zkj
rs("jiage")=jiage
rs("xin")=xin
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('产品增加成功，点击继续增加！');window.location.href='add_products.asp';</script>" 
end if
%>