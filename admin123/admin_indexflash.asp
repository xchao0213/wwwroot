<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(26)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>���ӻõ�</title>
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
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script> 
</head>
<body>
<%if request.querystring("action")="admin" then%>
<!--����Ϊ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��ҳflash�õƹ���</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox"/></td>
<td width="5%" align="center" class="td">ID</td>
<td width="22%" height="25" align="center" class="td">�� ��</td>
<td width="11%" align="center" class="td">����ͼ</td>
<td width="19%" align="center" class="td">������ɫ</td>
<td width="11%" align="center" class="td">����</td>
<td width="14%" align="center" class="td">ʱ��</td>
<td width="7%" align="center" class="td">�޸�</td>
<td width="8%" align="center" class="td">ɾ��</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from indexflash order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">���޼�¼!</div>")
else
rs.PageSize =300 'ÿҳ��¼����
iCount=rs.RecordCount '��¼����
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
<td width="5%" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="22%" height="25" align="center" class="td"><a href="xiugai_indexflash.asp?id=<%=rs("id")%>" style="color:#003399"><%=rs("title")%></a></td>
<td width="11%" align="center" class="td"><a href="<%=rs("img")%>"><img src="<%=rs("img")%>" width="80" height="28" border="0" /></a></td>
<td width="19%" align="center" class="td"><label>
  <input name="textfield" type="text" style="background:<%=rs("coler")%>; color:#FFFFFF" value="<%=rs("coler")%>" size="10" />
</label></td>
<td width="11%" align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td width="14%" align="center" class="td"><%=rs("tim")%></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit3" value="�޸�" onclick="window.location.href='admin_indexflash.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="8%" align="center" class="td">
<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">ȫѡ</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?action=del';" value="ɾ��"/>
<input type="submit" class="btn" onclick="form.action='?action=ordsc';" value="��������"/></td>
</tr>
</form>
</table>
</td>
  </tr>
</table>
<!--����Ϊ���-->
<div style="margin-top:10px">
<form  name="myform" method="post" action="?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��ҳflash�õ����</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
  <tr >
    <td width="16%" height="28" align="right" class="td">�������</td>
    <td width="84%"  class="td"><input name="title" type="text" size="30"  /></td>
  </tr>
  <tr bgcolor="#ffffff">
    <td width="16%" height="25" align="right" class="td"><font color="#FF0000">*</font>ͼƬ��ַ</td>
    <td colspan="2" class="td"><table width="81%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="28%"><input type="text" name="img" id="url" size="40" /></td>
<td width="10%"><input name="button2" type="button" id="filemanager" value="ѡȡ���ϴ�ͼƬ" class="btn" style="height:20px;" /></td>
<td width="62%"><iframe src="up.asp?formname=myform&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
</tr>
</table>
      </td>
    </tr>
  <tr>
    <td width="16%" height="25" align="right" class="td">���ӵ�ַ</td>
    <td class="td"><input name="link" type="text" value="#" size="30"  /> 
      ����������#��</td>
  </tr>
<tr>
    <td height="25" align="right" class="td"><font color="#FF0000">*</font>����ID</td>
    <td class="td"><input name="px_id" type="text" size="10"  /> 
      �������Ϊ���ֱ�����Ŀ</td>
  </tr>
    <tr>
    <td height="25" align="right" class="td"><font color="#FF0000">*</font>������ɫ</td>
    <td class="td"><input name="coler" type="text" size="10"  /> ��������ɫ</td>
  </tr>
  <tr>
    <td height="25" align="right" class="td">&nbsp;</td>
    <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
  </tr>

</table>
</td>
</tr>
</table>
</form>
<%end if%>
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from indexflash where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<form  name="myform" method="post" action="?xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸Ļõ�</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="16%" height="28" align="right" class="td">�������</td>
        <td  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="30"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="30"  /></td>
      </tr>
      <tr bgcolor="#ffffff">
        <td width="16%" height="25" align="right" class="td"><font color="#FF0000">*</font>ͼƬ��ַ</td>
        <td colspan="2" class="td"><table width="80%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%"><input type="text" name="img" id="url" size="40" value="<%=rs("img")%>"></td>
    <td width="11%"><input name="button2" type="button" id="filemanager" value="ѡȡ���ϴ�ͼƬ" class="btn" style="height:20px;" /></td>
    <td width="59%"><iframe src="up.asp?formname=myform&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="340" height="25"></iframe></td>
  </tr>
</table></td>
        </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">���ӵ�ַ</td>
        <td class="td"><input name="link" type="text" value="<%=rs("link")%>" size="30"  /> 
          ����������#��</td>
      </tr>
	  
    <tr>
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>����ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="10"  />
          �������Ϊ���ֱ�����Ŀ</td>
      </tr>
	<tr>
        <td height="25" align="right" class="td"><font color="#FF0000">*</font>������ɫ</td>
        <td class="td"><input name="coler" type="text" style="background:<%=rs("coler")%>; color:#FFFFFF" value="<%=rs("coler")%>"  size="10"  /> 
        ��������ɫ</td>
      </tr>	 
      <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="ȷ���޸�"  class="btn"/></td>
      </tr>
    
    </table>
</td>
  </tr>
</table></form>
<%end if%>
</div>
</body>
</html>
<% 
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from indexflash"
rs.open sql,conn,1,3
title=request.form("title")
link=request.form("link")
img=request.form("img")
px_id=request.form("px_id")
direction=request.form("direction")
slicing=request.form("slicing")
num=request.form("num")
shader=request.form("shader")
delay=request.form("delay")
coler=request.form("coler")
if title=""  then 
response.Write("<script language=javascript>alert('�����ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if link=""  then 
response.Write("<script language=javascript>alert('���ӵ�ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if img="" then 
response.Write("<script language=javascript>alert('ͼƬ��ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if px_id="" then 
response.Write("<script language=javascript>alert('����ID����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("link")=link
rs("img")=img
rs("px_id")=px_id
rs("direction")=direction
rs("slicing")=slicing
rs("num")=num
rs("shader")=shader
rs("delay")=delay
rs("coler")=coler
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('�����ӳɹ������������ӣ�');window.location.href='admin_indexflash.asp?action=admin';</script>" 
end if

if Request.QueryString("action")="ordsc" then
depname=request.form("px_id")
myarray=Split(depname,", ") 
set rs=server.createobject("adodb.recordset")
sql="select * from indexflash"
rs.open sql,conn,1,3
for i= 0 to Ubound(myarray)
rs("px_id")=myarray(i)
rs.update
rs.movenext
next
rs.close
set rs=nothing
Response.Write "<script>alert('������������ɹ�!');window.location.href='admin_indexflash.asp?action=admin';</script>" 
end if

if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from indexflash where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('ɾ���ɹ���');window.location.href='admin_indexflash.asp?action=admin';</script>"
end if
if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [indexflash] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_indexflash.asp?action=admin';</script>" 
end if
%>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
img=request.form("img")
link=request.form("link")
px_id=request.form("px_id")
direction=request.form("direction")
slicing=request.form("slicing")
num=request.form("num")
shader=request.form("shader")
delay=request.form("delay")
coler=request.form("coler")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from indexflash where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
if title=""  then 
response.Write("<script language=javascript>alert('�����ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if link=""  then 
response.Write("<script language=javascript>alert('���ӵ�ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if img="" then 
response.Write("<script language=javascript>alert('ͼƬ��ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if px_id="" then 
response.Write("<script language=javascript>alert('����ID����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
rs("title")=title
rs("img")=img
rs("link")=link
rs("px_id")=px_id
rs("direction")=direction
rs("slicing")=slicing
rs("num")=num
rs("shader")=shader
rs("delay")=delay
rs("coler")=coler
rs.update 
rs.close 
response.write "<script>alert('��ҳ�õ��޸ĳɹ���');window.location.href='admin_indexflash.asp?action=admin';</script>" 
end if
%>