<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(24)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>职位管理</title>
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
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">职位管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
  <td width="3%" align="center" class="td"><input name="ID" type="checkbox"/></td>
  <td width="4%" align="center" class="td">ID</td>
  <td width="21%" height="25" align="center" class="td">职位</td>
  <td width="17%" align="center" class="td">性别</td>
  <td width="10%" align="center" class="td">学历</td>
  <td width="11%" align="center" class="td">人数</td>
  <td width="20%" align="center" class="td">发布时间</td>
  <td width="7%" align="center" class="td">修改</td>
  <td width="7%" align="center" class="td">删除</td>
</tr></thead>
<form id="form" name="form" method="post" action="?">
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from job order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无记录!</div>")
else
rs.PageSize =30 '每页记录条数
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
for i=1 to rs.pagesize %>
<tr>
  <td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
  <td width="4%" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
  <td width="21%" height="25" align="center" class="td"><a href="admin_Job.asp?action=xiugai&id=<%=rs("id")%>" style="color:#003399"><%=rs("title")%></a> </td>
  <td width="17%" align="center" class="td"><%=rs("sex")%></td>
  <td width="10%" align="center" class="td"><%=rs("xueli")%></td>
  <td width="11%" align="center" class="td"><%=rs("renshu")%></td>
  <td width="20%" align="center" class="td"><%=rs("data")%></td>
  <td width="7%" align="center" class="td">
    <input type="button" name="Submit3" value="修改" onclick="window.location.href='admin_Job.asp?action=xiugai&id=<%=rs("id")%>'" class="btn"/></td>
  <td width="7%" align="center" class="td">
<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">全选</td>
<td colspan="8"><input type="submit" class="btn" onclick="form.action='?action=admin&act=del';" value="删除"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table></td>
  </tr>
</table>
<%end if%>
<%if request.querystring("action")="add" then%>
<form  name="add" method="post" action="admin_Job.asp?action=add&add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">发布招聘信息</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="16%" height="25" align="right" class="td">职位名称</td>
        <td width="84%"  class="td"><input name="title" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">年龄要求</td>
        <td class="td"><label>
          <input name="nn1" type="text" value="20"  size="10" /> -<input name="nn2" type="text" value="35" size="10" />岁</label></td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">性别要求</td>
        <td class="td">
          <input type="radio" name="sex"  value="男" />男
          <input type="radio" name="sex"  value="女" />女
          <input name="sex" type="radio"  value="不限" checked="checked" />不限</td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">要求学历</td>
        <td class="td">
          <select name="xueli" id="xueli">
            <option value="初中以上">初中以上</option>
            <option value="中专/高中">中专/高中</option>
            <option value="专科">专科</option>
            <option value="本科">本科</option>
            <option value="博士/硕士">博士/硕士</option>
            <option value="学历不限" selected="selected">学历不限</option>
          </select>        </td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">招聘人数</td>
      <td class="td"><input name="renshu" type="text"  size="10" /> 人</td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">其它要求</td>
        <td class="td"><label>
          <textarea name="body" cols="40" rows="10" style="width:700px;height:200px;visibility:hidden;"></textarea>
        </label></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
      </tr>
    </table>
    </td>
  </tr>
</table></form>
<%end if%>
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from job where id="&id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<form  name="add" method="post" action="?action=xiugai&id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改招聘信息</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="16%" height="25" align="right" class="td">职位名称</td>
        <td width="84%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="30"/>
          <input name="id" type="hidden" value="<%=rs("id")%>" size="30"/></td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">年龄要求</td>
        <td class="td"><label>
          <input name="nn1" type="text" value="<%=rs("nn1")%>"  size="10" /> -
        <input name="nn2" type="text" value="<%=rs("nn2")%>"  size="10" />岁</label></td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">性别要求</td>
        <td class="td">
<input type="radio" name="sex" value="男" <%if rs("sex")="男" then%>checked<%end if%>>男　 
<input type="radio" name="sex" value="女" <%if rs("sex")="女" then%>checked<%end if%>>女　 
<input type="radio" name="sex" value="不限" <%if rs("sex")="不限" then%>checked<%end if%>>不限  </td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">要求学历</td>
        <td class="td">
          <select name="xueli" id="xueli">
            <option value="初中以上" <%if rs("xueli")="初中以上" then%>selected<%end if%>>初中以上</option>
            <option value="中专/高中" <%if rs("xueli")="中专/高中" then%>selected<%end if%>>中专/高中</option>
            <option value="专科" <%if rs("xueli")="专科" then%>selected<%end if%>>专科</option>
            <option value="本科" <%if rs("xueli")="本科" then%>selected<%end if%>>本科</option>
            <option value="博士/硕士" <%if rs("xueli")="博士/硕士" then%>selected<%end if%>>博士/硕士</option>
            <option value="学历不限" <%if rs("xueli")="学历不限" then%>selected<%end if%>>学历不限</option>
          </select>       
           </td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">招聘人数</td>
      <td class="td"><input name="renshu" type="text" value="<%=rs("renshu")%>"  size="10" /> 
        人</td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">其它要求</td>
        <td class="td"><label>
          <textarea name="body" cols="40" rows="10" style="width:700px;height:200px;visibility:hidden;"><%=rs("body")%></textarea>
        </label></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
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
sql="select * from job where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('删除成功！');window.location.href='admin_job.asp?action=admin';</script>"
end if
if Request.QueryString("act")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from job where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_job.asp?action=admin';</script>" 
end if
%>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from job"
rs.open sql,conn,1,3
title=request.form("title")
body=request.form("body")
nn1=request.form("nn1")
nn2=request.form("nn2")
sex=request.form("sex")
xueli=request.form("xueli")
renshu=request.form("renshu")
if title=""  then 
response.Write("<script language=javascript>alert('职位名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if nn1="" or nn2=""  then 
response.Write("<script language=javascript>alert('年龄不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("nn1")) or not isNumeric(request("nn2")) then
response.write("<script>alert(""年龄必须为数字！""); history.go(-1);</script>")
response.end
end if
if renshu=""  then 
response.Write("<script language=javascript>alert('请填写招聘人数!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("renshu")) then
response.write("<script>alert(""招聘人数请填写数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("body")=body
rs("nn1")=nn1
rs("nn2")=nn2
rs("sex")=sex
rs("xueli")=xueli
rs("renshu")=renshu
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('职位发布成功！');window.location.href='admin_Job.asp?action=admin';</script>" 
end if
%>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
body=request.form("body")
nn1=request.form("nn1")
nn2=request.form("nn2")
sex=request.form("sex")
xueli=request.form("xueli")
renshu=request.form("renshu")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from job where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
if title=""  then 
response.Write("<script language=javascript>alert('职位名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if nn1="" or nn2=""  then 
response.Write("<script language=javascript>alert('年龄不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("nn1")) or not isNumeric(request("nn2")) then
response.write("<script>alert(""年龄必须为数字！""); history.go(-1);</script>")
response.end
end if
if renshu=""  then 
response.Write("<script language=javascript>alert('请填写招聘人数!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("renshu")) then
response.write("<script>alert(""招聘人数请填写数字！""); history.go(-1);</script>")
response.end
end if
rs("title")=title
rs("body")=body
rs("nn1")=nn1
rs("nn2")=nn2
rs("sex")=sex
rs("xueli")=xueli
rs("renshu")=renshu
rs.update 
rs.close 
response.write "<script>alert('修改成功！');window.location.href='admin_Job.asp?action=admin';</script>" 
end if
%> 