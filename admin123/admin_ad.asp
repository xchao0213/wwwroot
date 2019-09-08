<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(25)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>广告管理</title>
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
 
function oCopy(obj){ 
obj.select(); 
js=obj.createTextRange(); 
js.execCommand("Copy") 
alert('调用代码已复到剪贴板！请粘贴到前台需要显示广告的地方即可!')
} 

<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script> 
</head>
<body>
<%if request.querystring("action")="add" then%>
<form  name="myform" method="post" action="admin_ad.asp?add=ok">
<%id=Request.QueryString("id")%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加广告</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
<%if id=3 then%>
<tr >
<td width="16%" height="28" align="right" class="td">广告名称 <font color="#FF0000">*</font></td>
<td width="84%"  class="td">
<input name="title" type="text" size="40"  /></td>
</tr>
<tr>
<td width="16%" height="25" align="right" class="td">广告类型 <font color="#FF0000">*</font></td>
<td class="td">     

<select name="lxss" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option value="admin_ad.asp?action=add&id=1" <%if id=1 then%>selected="selected"<%end if%>>图片广告</option>
<option value="admin_ad.asp?action=add&id=2" <%if id=2 then%>selected="selected"<%end if%>>FLASH动画</option>
<option value="admin_ad.asp?action=add&id=3" <%if id=3 then%>selected="selected"<%end if%>>文字广告</option>
</select>
<label>
<input name="lx" type="hidden" id="lx" value="<%= request.QueryString("id")%> " />
</label></td>
</tr>
      <tr >
<td width="16%" height="28" align="right" class="td">文字内容 <font color="#FF0000">*</font></td>
<td width="84%"  class="td">
<textarea name="code" cols="40" rows="10"></textarea></td>
</tr>
<%else%>
<tr >
<td width="16%" height="28" align="right" class="td">广告名称 <font color="#FF0000">*</font></td>
<td width="84%"  class="td">
<input name="title" type="text" size="40"  /></td>
</tr>
<tr>
<td width="16%" height="25" align="right" class="td">广告类型 <font color="#FF0000">*</font></td>
<td class="td">     

<select name="lxss" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option value="admin_ad.asp?action=add&id=1" <%if id=1 then%>selected="selected"<%end if%>>图片广告</option>
<option value="admin_ad.asp?action=add&id=2" <%if id=2 then%>selected="selected"<%end if%>>FLASH动画</option>
<option value="admin_ad.asp?action=add&id=3" <%if id=3 then%>selected="selected"<%end if%>>文字广告</option>
</select>
<input name="lx" type="hidden" id="lx" value="<%= request.QueryString("id")%> " /></td>
</tr>

<tr bgcolor="#ffffff">
<td width="16%" height="25" align="right" class="td">图片地址 <font color="#FF0000">*</font></td>
<td colspan="2" class="td"><table width="70%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="35%"><input type="text" name="img" size="40" /></td>
<td width="65%"><iframe src="up.asp?formname=myform&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="350" height="21"></iframe></td>
</tr>
</table>
</td>
</tr>
<tr>
<td height="13" align="right" class="td">链接地址 <font color="#FF0000">*</font></td>
<td class="td"><input name="url" type="text" value="http://" size="40"  /> <label>
打开方式：
  <select name="openfs" id="openfs">
<option value="_blank">_blank</option>
<option value="_parent">_parent</option>
<option value="_self">_self</option>
<option value="_top">_top</option>
</select>
</label></td>
</tr>
<tr>
<td height="6" align="right" class="td">广告尺寸</td>
<td class="td"><input name="width" type="text" size="10"  />
<font color="#FF0000">*</font>
宽
<input name="height" type="text" size="10"  />
<font color="#FF0000">*</font>          高</td>
</tr>
<tr>
<td height="6" align="right" class="td">链接说明</td>
<td class="td"><input name="sm" type="text" size="40"  /></td>
</tr>
<tr>
<td height="25" align="right" class="td"><p>备注</p>
</td>
<td class="td"><textarea name="bz" cols="50" rows="5"></textarea></td>
</tr><%end if%>
<tr>
<td height="25" align="right" class="td">&nbsp;</td>
<td class="td"><input type="submit" name="button" id="button" value="确认提交"  class="btn"/></td>
</tr>
</table>
</td>
  </tr>
</table></form>
<%end if
if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">广告管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="3%" align="center" class="td">ID</td>
<td width="26%" height="25" align="center" class="td">广告名称</td>
<td width="9%" align="center" class="td">缩略图</td>
<td width="28%" align="center" class="td">调用代码</td>
<td width="8%" align="center" class="td">点击</td>
<td width="8%" align="center" class="td">生成</td>
<td width="7%" align="center" class="td">修改</td>
<td width="8%" align="center" class="td">删除</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from ad order by id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无记录!</div>")
else
rs.PageSize =20 '每页记录条数
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
<td width="3%" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="26%" height="25" align="center" class="td"><a href="admin_ad.asp?action=xiugai&id=<%=rs("id")%>" style="color:#003399"><%=rs("title")%></a> </td>
<td width="9%" align="center" class="td"><%
if IsNull(rs("img")) or rs("img")="" then
else
response.write ("<a href="&rs("img")&"><img src="&rs("img")&" width=""70"" height=""25"" alt=""图片缩略图""/></a>")
end if%>
</td>
<td width="28%" align="center" class="td"><input onclick="oCopy(this)" value="&lt;script src=/js/zychad<%=rs("id")%>.js&gt;&lt;/script&gt;" size="30"/></td>
<td width="8%" align="center" class="td"><%=rs("hit")%></td>
<td width="8%" align="center" class="td"><input type="button" name="Submit2" value="生成JS" onclick="window.location.href='admin_ad.asp?action=admin&id=<%=rs("id")%>&Generation=ok' "  class="btn"/>
</td>
<td width="7%" align="center" class="td"><input type="button" name="Submit3" value="修改" onclick="window.location.href='admin_ad.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="8%" align="center" class="td">
<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除当前广告？删除后不可恢复!')){window.location.href='admin_ad.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="3%" height="30" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="3%" height="30" align="center" class="td">全选</td>
<td colspan="7" class="td"><input type="submit" class="btn" onclick="form.action='?action=admin&act=del';" value="删除"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr></form>
<tr><td colspan="9">&nbsp;注意：增加广告或修改广告后勿必生成JS再调用！</td></tr>
</table></td>
  </tr>
</table>
<!--以下为修改-->
<%end if
if request.querystring("action")="xiugai" then
exec="select * from ad where id="& request.QueryString("id") 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
id=rs("id")
lx=rs("lx")
%>
<form  name="add" method="post" action="admin_ad.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">修改广告</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
          <tr >
        <td width="16%" height="14" align="right" class="td"><font color="#FF0000">*</font>广告类型</td>
        <td width="84%" height="25"  class="td"><strong>
		<%if lx=1 then
		response.Write("图片广告")
		elseif lx=2 then
		response.Write("FLASH广告")
		else
		response.Write("文字广告")
		end if
		%></strong>
          <input name="lx" type="hidden" id="lx" value="<%=rs("lx")%>" /></td>
          </tr>
       <%if lx=3 then%>  
      <tr >
        <td width="16%" height="28" align="right" class="td"><font color="#FF0000">*</font>广告名称</td>
        <td width="84%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="40"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="40"  /></td>
      </tr> 
<tr >
                          <td height="14" align="right" class="td"><font color="#FF0000">*</font>文字内容</td>
<td width="84%"  class="td"><textarea name="code" cols="40" rows="10"><%=rs("code")%></textarea></td>
                        </tr><%else%>
         <tr >
        <td width="16%" height="28" align="right" class="td"><font color="#FF0000">*</font>广告名称</td>
        <td width="84%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="40"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="40"  /></td>
      </tr> 
      <tr bgcolor="#ffffff">
        <td width="16%" height="25" align="right" class="td"><font color="#FF0000">*</font>图片地址</td>
        <td colspan="2" class="td"><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="32%"><input type="text" name="img" id="url" size="40" value="<%=rs("img")%>"></td>
    <td width="9%"><input name="button2" type="button" id="filemanager" value="选取已上传" /></td>
    <td width="59%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="350" height="21"></iframe></td>
  </tr>
</table> 
</td>
  </tr>
    <tr>
        <td height="13" align="right" class="td"><font color="#FF0000">*</font>链接地址</td>
        <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="40"  />
          打开方式：
          <select name="openfs" id="openfs">
            <option value="_blank"  <%if rs("openfs")="_blank" then%>selected="selected"<%end if%>>_blank</option>
            <option value="_parent" <%if rs("openfs")="_parent" then%>selected="selected"<%end if%>>_parent</option>
            <option value="_self" <%if rs("openfs")="_self" then%>selected="selected"<%end if%>>_self</option>
            <option value="_top" <%if rs("openfs")="_top" then%>selected="selected"<%end if%>>_top</option>
          </select></td>
    </tr>
    <tr>
      <td height="6" align="right" class="td">广告尺寸</td>
      <td class="td"><input name="width" type="text" value="<%=rs("width")%>" size="10"  />
        <font color="#FF0000">*</font>
        宽
          <input name="height" type="text" value="<%=rs("height")%>" size="10"  />
          <font color="#FF0000">*</font>          高</td>
    </tr>
    <tr>
      <td height="6" align="right" class="td">链接说明</td>
      <td class="td"><input name="sm" type="text" value="<%=rs("sm")%>" size="40"  /></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td"><p>备注</p>          </td>
        <td class="td"><textarea name="bz" cols="50" rows="5"><%=rs("bz")%></textarea></td>
      </tr>
    <tr>
        <td height="13" align="right" class="td">发布时间</td>
        <td class="td"><input name="data" type="text" value="<%=rs("data")%>" size="40"   onfocus="show_cele_date(data,'','',data)"/></td>
      </tr>  <%end if%>
      <tr>
      <td height="12" align="right" class="td">&nbsp;</td>
      <td class="td"><input type="submit" name="button" id="button" value="提交修改"  class="btn"/></td>
    </tr>
  
    </table>
</td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from ad"
rs.open sql,conn,1,3
title=request.form("title")
url=request.form("url")
width=request.form("width")
height=request.form("height")
img=request.form("img")
lx=request.form("lx")
sm=request.form("sm")
bz=request.form("bz")
openfs=request.form("openfs")
code=request.form("code")
if title=""  then 
response.Write("<script language=javascript>alert('广告名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if lx=1 or lx=2  then
if img=""  then 
response.Write("<script language=javascript>alert('广告图片不能为空!');history.go(-1)</script>") 
response.end 
end if
if url=""  then 
response.Write("<script language=javascript>alert('广告链接地址不能为空!');history.go(-1)</script>") 
response.end 
end if
if width=""  then 
response.Write("<script language=javascript>alert('请输入广告宽度!');history.go(-1)</script>") 
response.end 
end if
if height=""  then 
response.Write("<script language=javascript>alert('请输入广告高度!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("width")) or not isNumeric(request("height"))then
response.write("<script>alert(""广告尺寸必须都为数字！""); history.go(-1);</script>")
response.end
end if
else
if code=""  then 
response.Write("<script language=javascript>alert('文字广告内容不能为空!');history.go(-1)</script>") 
response.end 
end if
end if
rs.addnew
rs("title")=title
rs("url")=url
rs("width")=width
rs("height")=height
rs("img")=img
rs("lx")=lx
rs("sm")=sm
rs("bz")=bz
rs("code")=code
rs("openfs")=openfs
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('广告增加成功了哈！');window.location.href='admin_ad.asp?action=admin';</script>" 
end if
%>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from ad where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('当前广告删除成功！');window.location.href='admin_ad.asp?action=admin';</script>"
end if
if Request.QueryString("act")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [ad] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_ad.asp?action=admin';</script>" 
end if

if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
lx=request.form("lx")
img=request.form("img")
url=request.form("url")
width=request.form("width")
height=request.form("height")
sm=request.form("sm")
bz=request.form("bz")
code=request.form("code")
data=request.form("data")
openfs=request.form("openfs")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from ad where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
	Response.End()
	end if
	if title=""  then 
if lx=1 or lx=2  then
response.Write("<script language=javascript>alert('广告名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if img=""  then 
response.Write("<script language=javascript>alert('广告图片不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("width")) or not isNumeric(request("height"))then
response.write("<script>alert(""广告尺寸必须都为数字！""); history.go(-1);</script>")
response.end
end if
elseif lx=3 then
if code=""  then 
response.Write("<script language=javascript>alert('文字内容不能为空!');history.go(-1)</script>") 
response.end 
end if
else
end if
rs("title")=title
rs("url")=url
rs("width")=width
rs("height")=height
rs("img")=img
rs("lx")=lx
rs("sm")=sm
rs("bz")=bz
rs("openfs")=openfs
rs("data")=data
rs("code")=code
rs.update 
rs.close 
set js = server.CreateObject("ADODB.RecordSet")
sql="select * from ad where id="&request.QueryString("id")
set js = conn.Execute (Sql) 
if js("lx")=1 then
goaler = goaler + "<a href="""&dir&"Include/Showbody.asp?action=adurl&id="& js("id")&""" target="""& js("openfs")&"""><img src="""& js("img")&""" width="""& js("width")&""" height="""& js("height")&"""  title="""& js("sm")&"""></a>"  
elseif js("lx")=2 then
goaler = goaler + "<embed src="""&js("img")&""" quality=""height"" type=""application/x-shockwave-flash""  width="""& js("width")&""" height="""& js("height")&""" ></embed>" 
else
goaler = goaler + ""& js("code")&"" 
end if
'生成JS文件
goaler = "" + goaler + ""
goaler = "document.write('" & goaler & "')"
FolderPath = Server.MapPath("../")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.CreateTextFile(FolderPath&"/js/zychad"&js("id")&".js")
fout.WriteLine goaler
'关闭连接
fout.close
set fout = nothing
js.close
set js = nothing
conn.close
set conn=nothing
response.write "<script>alert('广告修改成功！');window.location.href='admin_ad.asp?action=admin';</script>" 
end if
%> 
<%
if Request.QueryString("Generation")="ok" then
set js = server.CreateObject("ADODB.RecordSet")
sql="select * from ad where id="&request.QueryString("id")
set js = conn.Execute (Sql) 
if js("lx")=1 then
goaler = goaler + "<a href="""&dir&"Include/Showbody.asp?action=adurl&id="& js("id")&""" target="""& js("openfs")&"""><img src="""& js("img")&""" width="""& js("width")&""" height="""& js("height")&"""  title="""& js("sm")&"""></a>"  
elseif js("lx")=2 then
goaler = goaler + "<embed src="""&js("img")&""" quality=""height"" type=""application/x-shockwave-flash""  width="""& js("width")&""" height="""& js("height")&""" ></embed>" 
else
goaler = goaler + ""& js("code")&"" 
end if
'生成JS文件
goaler = "" + goaler + ""
goaler = "document.write('" & goaler & "')"
FolderPath = Server.MapPath("../")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.CreateTextFile(FolderPath&"/js/zychad"&js("id")&".js")
fout.WriteLine goaler
'关闭连接
fout.close
set fout = nothing
js.close
set js = nothing
conn.close
set conn=nothing
Response.Write "<script>alert('广告JS已经生成!');window.location.href='admin_ad.asp?action=admin';</script>" 
end if
%>