<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(36)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>�������ӹ���</title>
<link rel="stylesheet" href="../zycheditor/themes/default/default.css" />
<link rel="stylesheet" href="../zycheditor/plugins/code/prettify.css" />
<script charset="utf-8" src="../zycheditor/kindeditor.js"></script>
<script charset="utf-8" src="../zycheditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zycheditor/plugins/code/prettify.js"></script>
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
              K('#logo').val(url);
              editor.hideDialog();
          }
      });
  });
});
});
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
<%if request.querystring("action")="img" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">ͼƬ���ӹ���</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox"/></td>
<td width="3%" align="center" class="td">ID</td>
<td width="19%" align="center" class="td">LOGO</td>
<td width="24%" height="25" align="center" class="td">�� ��</td>
<td width="29%" align="center" class="td">���ӵ�ַ</td>
<td width="10%" align="center" class="td">����</td>
<td width="6%" align="center" class="td">�޸�</td>
<td width="6%" align="center" class="td">ɾ��</td>
</tr></thead>
<form action="?" method="post" name="form1">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from link where lx=2 order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.Write "������Ϣ��"
else
rs.PageSize =500 'ÿҳ��¼����
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
for i=1 to rs.pagesize%>
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="3%" align="center" class="td"><input type="text" readonly="readonly" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="19%" align="center" class="td"><img src="<%=rs("logo")%>" width="88" height="31" /></td>
<td width="24%" height="25" align="center" class="td"><%=rs("title")%></td>
<td width="29%" align="center" class="td"><a href="<%=rs("url")%>" target="_blank"><%=rs("url")%></a></td>
<td width="10%" align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td width="6%" align="center" class="td"><input type="button" name="Submit3" value="�޸�" onclick="window.location.href='admin_link.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="6%" align="center" class="td"><input type="button" name="Submit" value="ɾ��" onClick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='?act=del&id=<%=rs("id")%>';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">ȫѡ</td>
<td colspan="4"><input type="submit" class="btn" onclick="form.action='?action=img&del=ok';" value="ɾ��"/>
<input type="submit" class="btn" onclick="form.action='?action=img&xiugai=img';" value="��������"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>            
</table></td>
  </tr>
</table>
<% End If %>
<%if request.querystring("action")="txt" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�������ӹ���</div></td>
  </tr>
  <tr>
<td height="35" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox"/></td>
<td width="3%" align="center" class="td">ID</td>
<td height="25" align="center" class="td">��������</td>
<td width="29%" align="center" class="td">���ӵ�ַ</td>
<td align="center" class="td">����ID</td>
<td width="9%" align="center" class="td">ɾ��</td>
</tr></thead>
<form action="?" method="post" name="form1">
<%
set rs=server.createobject("adodb.recordset") 
exec="select * from link where lx=1 order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.Write "&nbsp;������Ϣ��"
else
rs.PageSize =500 'ÿҳ��¼����
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
for i=1 to rs.pagesize%>
<tr>
<td align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td align="center" class="td"><input type="text" readonly="readonly" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td height="25" align="center" class="td"><input name="title" type="text" value="<%=rs("title")%>" size="18" /></td>
<td align="center" class="td"><input name="url" type="text" value="<%=rs("url")%>" size="25" /></td>
<td align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td align="center" class="td"><input type="button" name="Submit" value="ɾ��" onClick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='?act=del&id=<%=rs("id")%>';}else{history.go(0);}"  class="btn"/></td>
</tr>
<% rs.movenext 
if rs.eof then exit for 
next 
end if%> 
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">ȫѡ</td>
<td colspan="4"><input type="submit" class="btn" onclick="form.action='?action=txt&del=ok';" value="ɾ��"/>
<input type="submit" class="btn" onclick="form.action='?action=txt&xiugai=txt';" value="�����޸�"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr></form>       
</table>
</td>
  </tr>
</table>
<% End If %>
<%if request.querystring("action")="add" then%>
<%lx=Request.QueryString("lx")%>
<form  name="myform" method="post" action="admin_link.asp?action=add&lx=<%=lx%>&add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">������������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="15%" height="13" align="right" class="td">ѡ������</td>
        <td width="85%"  class="td"><select name="lxxz" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
 <option value="?action=add&lx=1" <%if lx=1 then%>selected="selected"<%end if%>>��������</option>
<option value="?action=add&lx=2" <%if lx=2 then%>selected="selected"<%end if%>>ͼƬ����</option>
</select>
          <label>
          <input name="lx" type="hidden" value="<%=lx%>" />
          </label></td>
      </tr>
      <tr >
        <td width="15%" height="12" align="right" class="td">��������</td>
        <td  class="td"><input name="title" type="text" size="30"  /></td>
      </tr>
     <%if lx="" or lx=1 then %>
      <tr>
        <td width="15%" height="25" align="right" class="td">���ӵ�ַ</td>
        <td class="td"><input name="url" type="text" value="http://" size="30"  /></td>
      </tr>
      <tr>
        <td width="15%" height="25" align="right" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /></td>
      </tr>
    <%else%>
      <tr bgcolor="#FFFFFF">
        <td height="13" align="right" class="td">LOGOͼ��</td>
        <td class="td"><table width="78%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%"><input type="text" name="logo" id="logo" size="40" /></td>
    <td width="12%"><input name="button" type="button" id="filemanager" value="ѡȡ���ϴ�ͼƬ" class="btn" style="height:20px;" /></td>
    <td width="55%"><iframe src="up.asp?formname=myform&inputname=logo&uploadstyle=Logo" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
  </tr>
</table></td>
      </tr>
      <tr>
        <td height="6" align="right" class="td">���ӵ�ַ</td>
        <td class="td"><input name="url" type="text" value="http://" size="30"  /></td>
      </tr>
      <tr>
        <td height="6" align="right" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /></td>
      </tr>
 <%end if%>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
      </tr>
    </table></td>
  </tr>
</table></form>
<% end if %>
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from link where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if
%>
<form  name="myform" method="post" action="admin_link.asp?action=xiugai&id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸�ͼƬ����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="15%" height="12" align="right" class="td">��������</td>
        <td width="85%"  class="td"><input name="title" type="text" value="<%=rs("title")%>" size="30"  /></td>
      </tr>
      <tr>
        <td width="15%" height="25" align="right" class="td">���ӵ�ַ</td>
        <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="30"  /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="13" align="right" class="td">LOGOͼ��</td>
        <td class="td"><table width="63%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="32%"><input name="logo" type="text" id="logo" value="<%=rs("logo")%>" size="40" /></td>
    <td width="11%"><input name="button" type="button" id="filemanager" value="ѡȡ���ϴ�ͼƬ" class="btn" style="height:20px;" /></td>
    <td width="57%"><iframe src="up.asp?formname=myform&inputname=logo&uploadstyle=Logo" frameborder="0" scrolling="no" width="350" height="25"></iframe></td>
  </tr>
</table></td>
      </tr>
      <tr>
        <td height="6" align="right" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="30"  /></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/>
		<input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
    </table></td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<% '====�޸���������
if Request.QueryString("xiugai")="txt" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
id=Split(request.form("id"),", ")
title=Split(request.form("title"),", ")
url=Split(request.form("url"),", ")  
px_id=Split(request.form("px_id"),", ") 
set rs=server.createobject("adodb.recordset")
sql="select * from link"
rs.open sql,conn,1,3
for i= 0 to Ubound(id)
rs("title")=title(i)
rs("url")=url(i)
rs("px_id")=px_id(i)
rs.update
rs.movenext
next
rs.close
set rs=nothing
response.Write("<script language=""javascript"">alert(""�޸����ӳɹ�"");window.location.href='admin_link.asp?action=txt';</script>")
end if
'==����ɾ��
if Request.QueryString("del")="ok" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from link where id in ("&Request("id")&")"
conn.execute(sql) 
	if Request.QueryString("action")="img" then
	Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_link.asp?action=img';</script>" 
	else
	Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_link.asp?action=txt';</script>" 
	end if
end if

if request("act")="del" then
	id=request("id")
	if id="" then
	response.Write("<script language=javascript>alert('��������!');history.go(-1)</script>") 
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from Link where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	response.Write("<script language=javascript>alert('���ݿ���û�иü�¼!');history.go(-1)</script>") 
	Response.End()
else
	rs.Delete
	rs.Update
	if Request.QueryString("action")="img" then
	Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_link.asp?action=img';</script>" 
	else
	Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_link.asp?action=txt';</script>" 
	end if
end if
end if
'������
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from link"
rs.open sql,conn,1,3
title=request.form("title")
url=request.form("url")
px_id=request.form("px_id")
lx=request.form("lx")
logo=request.form("logo")
if title=""  then 
response.Write("<script language=javascript>alert('�������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if url=""  then 
response.Write("<script language=javascript>alert('���ӵ�ַ����Ϊ��!');history.go(-1)</script>") 
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
if lx="" then
lx=1
end if
rs.addnew
rs("lx")=lx
rs("title")=title
rs("url")=url
rs("px_id")=px_id
rs("logo")=logo
rs.update
rs.close
set rs=nothing
if Request.QueryString("lx")=2 then
Response.Write "<script>alert('ͼƬ�����������ӳɹ�������������ӣ�');window.location.href='admin_link.asp?action=img';</script>" 
else
Response.Write "<script>alert('���������������ӳɹ�������������ӣ�');window.location.href='admin_link.asp?action=txt';</script>" 
end if
end if%>
<% 
if request("xiugai")="ok" then
id=request("id")
title=request.form("title")
url=request.form("url")
logo=request.form("logo")
px_id=request.form("px_id")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from link where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
	if title=""  then 
response.Write("<script language=javascript>alert('�������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('����ID����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
rs("title")=title
rs("url")=url
rs("logo")=logo
rs("px_id")=px_id
rs.update 
rs.close 
response.write "<script>alert('�޸ĳɹ���');window.location.href='admin_link.asp?action=img';</script>"
end if 
%>