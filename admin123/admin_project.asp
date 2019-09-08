<!--#include file="../Include/Conn.asp" -->
<!--#include file="page.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(17)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<script src="images/color.js" type="text/javascript"></script>
<title>服务项目管理</title>
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
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form.elements.length);i++){ 
  var e=form.elements[i]; 
  if (e.name!="allbox") e.checked=form.allbox.checked; 
 } 
} 
--> 
</script> 
</head>
<body>
<%if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">服务项目管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
  <td width="20" align="center" class="td"><input type="checkbox" /></td>
  <td width="40" align="center" class="td">ID</td>
  <td width="102" height="41" align="center" class="td">项目名称</td>
  <td width="158" align="center" class="td">图 片</td>
  <td width="391" align="center" class="td">说 明</td>
  <td width="104" align="center" class="td">排 序</td>
  <td width="92" align="center" class="td">修 改</td>
  <td width="93" align="center" class="td">删 除</td>
</tr></thead>
<form id="form" name="form" method="post" action="?"> 
<%
set rs=server.createobject("adodb.recordset") 
exec="select * from project order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无服务项目!</div>")
else
rs.PageSize =300 '每页记录条数
iCount=rs.RecordCount '记录总数
iPageSize=rs.PageSize
maxpage=rs.PageCount 
page=request("page")
if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
if page=maxpage then
x=iCount-(maxpage-1)*iPageSize
else
x=iPageSize
end if	
for i=1 to rs.pagesize%>
<tr>
  <td width="20" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
  <td width="40" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
  <td width="102" height="41" align="center" class="td"><%=rs("name")%></td>
  <td width="158" align="center" class="td"><img src="<%=rs("img")%>" width="70" height="30" border="0"/></td>
  <td width="391" class="td"><%=stvalue(DelHtml(rs("shuomin")),50)%></td>
  <td width="104" align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
  <td width="92" align="center" class="td"><input type="button" name="Submit3" value="修改" onclick="window.location.href='admin_project.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
  <td width="93" align="center" class="td"><input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_project.asp?action=admin&id=<%=rs("id")%>&del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="20" height="30" align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="40" height="30" align="center">全选</td>
<td height="30" colspan="8"><input type="submit" class="btn" onclick="form.action='?action=del';" value="删除"/>
<input type="submit" class="btn" onclick="form.action='?action=ordsc';" value="更新排序"/>
</label><%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>
</td>
  </tr>
</table>
<!--以下为增加内容-->
<%end if
if request.querystring("action")="add" then%>
<form action="admin_project.asp?add=ok" method="post" name="add">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
 <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改服务</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr  >
        <td width="13%" height="28" align="right" class="td"><font color="#FF0000">*</font>服务标题</td>
        <td width="87%"  class="td"><input name="name" type="text" value="" size="40"  /></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td width="13%" height="13" align="right" class="td"><font color="#FF0000">*</font>排序ID</td>
        <td class="td"><input name="px_id" type="text" value="" size="40"></td>
      </tr>
     
<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
<td height="12" align="right" class="td"><font color="#FF0000">*</font>标题略缩图</td>
<td class="td"><table width="52%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="22%"><input type=text id="url" name=img size=40  value=""></td>
<td width="8%"><input name="button2" type="button" id="filemanager" class="btn2" value="选取已上传图片" /></td>
<td width="70%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=case" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
</tr>
</table></td>
</tr>
<tr >
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>简单说明</td>
        <td class="td"><textarea name="shuomin" cols="50" rows="4"></textarea></td>
      </tr>

     <tr >
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>服务详情</td>
        <td class="td"><textarea name="body" cols="50" style="width:700px;height:200px;visibility:hidden;"></textarea></td>
      </tr>
       <tr>
          <td height="12" align="right" class="td">&nbsp;</td>
          <td class="td"><input type="submit" name="button" id="button" value="确认修改"  class="btn"/></td>
        </tr>
    </table>
</td>
  </tr>
</table></form>
<!--以下为修改内容-->
<%end if
if request.querystring("action")="xiugai" then 
exec="select * from project where id="& request.QueryString("id") 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<form  name="add" method="post" action="admin_project.asp?action=xiugai&id=<%=rs("id")%>&xiugai=ok"> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
 <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改服务</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr  >
        <td width="13%" height="28" align="right" class="td"><font color="#FF0000">*</font>服务标题</td>
        <td width="87%"  class="td">
          <input name="name" type="text" value="<%=rs("name")%>" size="40"  />
          <label><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /> </label></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td width="13%" height="13" align="right" class="td"><font color="#FF0000">*</font>排序ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="40"></td>
      </tr>
     
<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
<td height="12" align="right" class="td"><font color="#FF0000">*</font>标题略缩图</td>
<td class="td"><table width="52%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="22%"><input type=text id="url" name=img size=40  value="<%=rs("img")%>"></td>
<td width="8%"><input name="button2" type="button" id="filemanager" class="btn2" value="选取已上传图片" /></td>
<td width="70%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=case" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
</tr>
</table></td>
</tr>
<tr >
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>简单说明</td>
        <td class="td"><textarea name="shuomin" cols="50" rows="4"><%=rs("shuomin")%></textarea></td>
      </tr>

     <tr >
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>服务详情</td>
        <td class="td"><textarea name="body" cols="50" style="width:700px;height:200px;visibility:hidden;"><%=rs("body")%></textarea></td>
      </tr>
       <tr >
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>原版面图片</td>
        <td colspan="2" class="td">
<%
if IsNull(rs("img")) or rs("img")="" then
else
response.write ("<img src="&rs("img")&" width=""200px"" alt=""原版面图片""/>")
end if
%></td>
	  </tr>

       <tr >
          <td height="12" align="right" class="td">&nbsp;</td>
          <td class="td"><input type="submit" name="button" id="button" value="确认修改"  class="btn"/></td>
        </tr>
    </table>
</td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from project where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('当前照片删除成功！');window.location.href='admin_project.asp?action=admin';</script>"
end if 
%>
<%
if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from project where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_project.asp?action=admin';</script>" 
end if
'=====================================

if Request.QueryString("action")="ordsc" then
depname=request.form("px_id")
myarray=Split(depname,", ") 
set rs=server.createobject("adodb.recordset")
sql="select * from project order by px_id asc"
rs.open sql,conn,1,3
for i= 0 to Ubound(myarray)
rs("px_id")=myarray(i)
rs.update
rs.movenext
next
rs.close
set rs=nothing
Response.Write "<script>alert('排序成功！');window.location.href='admin_project.asp?action=admin';</script>"
end if

if request("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from project"
rs.open sql,conn,1,3
name=request.form("name")
px_id=request.form("px_id")
img=request.form("img")
shuomin=request.form("shuomin")
body=request.form("body")
if name=""  then 
response.Write("<script language=javascript>alert('项目名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if img=""  then 
response.Write("<script language=javascript>alert('图片不能为空!');history.go(-1)</script>") 
response.end 
end if
if shuomin=""  then 
response.Write("<script language=javascript>alert('说明不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert(""错误！排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("name")=name
rs("img")=img
rs("shuomin")=shuomin
rs("body")=body
rs("px_id")=px_id
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('项目增加成功！');window.location.href='admin_project.asp?action=admin';</script>" 
end if

if request("xiugai")="ok" then
id=request("id")
name=request.form("name")
px_id=request.form("px_id")
img=request.form("img")
shuomin=request.form("shuomin")
body=request.form("body")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from project where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if
if name=""  then 
response.Write("<script language=javascript>alert('服务名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
if shuomin=""  then 
response.Write("<script language=javascript>alert('简单说明不能为空!');history.go(-1)</script>") 
response.end 
end if
rs("name")=name
rs("img")=img
rs("shuomin")=shuomin
rs("body")=body
rs("px_id")=px_id
rs.update 
rs.close 
response.write "<script>alert('服务项目修改成功！');window.location.href='admin_project.asp?action=admin';</script>" 
end if
%> 
