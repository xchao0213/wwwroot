<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(10)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>�Ļ��������</title>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�Ļ��������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="6%" align="center" class="td">ID</td>
<td width="15%" height="25" align="center" class="td">��������</td>
<td width="23%" align="center" class="td">�ؼ���</td>
<td width="29%" align="center" class="td">�� ַ</td>
<td width="13%" align="center" class="td">����</td>
<td width="7%" align="center" class="td">�޸�</td>
<td width="7%" align="center" class="td">ɾ��</td>
</tr></thead>
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from culture_fl order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">�����Ļ�����!</div>")
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
<form action="admin_culture_fl.asp?xiugai=ok" method="post" name="add">
<tr>
<td width="6%" align="center" class="td"><input name="id" type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="15%" height="25" align="center" class="td"><input name="title" type="text" size="15"  value="<%=rs("title")%>"/></td>
<td width="23%" align="center" class="td"><input name="key" type="text" size="20"  value="<%=rs("key")%>"/></td>
<td width="29%" align="center" class="td"><input name="url" type="text" value="/Culture/Culturelist.asp?id=<%=rs("id")%>" size="20"/> 
[<a href="/Culture/Culturelist.asp?id=<%=rs("id")%>" target="_blank" style="color:#003399">����</a>]</td>
<td width="13%" align="center" class="td"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td width="7%" align="center" class="td"><input type="submit" name="button2" id="button2" value="�޸�"  class="btn"/></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit" value="ɾ��" onClick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='admin_culture_fl.asp?act=del&id=<%=rs("id")%>';}else{history.go(0);}"  class="btn"/></td>
</tr></form>
<% rs.movenext 
if rs.eof then exit for 
next 
end if%>
</table></td>
  </tr>
</table>

<div style="margin-top:10px">

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="admin_culture_fl.asp?add=ok" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">���ӷ���</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="16%" height="25" align="right" class="td">��������</td>
        <td width="84%"  class="td"><input name="title" type="text" size="30"/></td>
      </tr>
      
      <tr>
        <td width="16%" height="25" align="right" class="td">�ؼ���</td>
        <td class="td"><input name="key" type="text" size="30"  /> ����ؼ�������,����</td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /> ����ԽСԽ��ǰ��</td>
      </tr>
      
      <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><label>
        <input type="submit" name="button" id="button" value="���ӷ���"  class="btn"/>
        </label></td>
      </tr>
    </table></td>
  </tr></form>
</table>
</div>
</body>
</html>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from culture_fl"
rs.open sql,conn,1,3
title=request.form("title")
px_id=request.form("px_id")
if title=""  then 
response.Write("<script language=javascript>alert('�������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('����ID����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert("" ��������ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("px_id")=px_id
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('�������ӳɹ���');window.location.href='admin_culture_fl.asp';</script>" 
end if
%>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id") 
sql="select * from culture_fl where id="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
rs("title")=request.form("title") 
rs("key")=request.form("key") 
rs("px_id")=request.form("px_id") 
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""��ǰ�����޸ĳɹ���"");window.location.href='admin_culture_fl.asp';</script>")
end if
%> 
<%
if request("act")="del" then
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('admin_culture_fl.asp');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from culture_fl where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('���ݿ���û�иü�¼��');document.location.href('admin_culture_fl.asp');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('��ǰ����ɾ���ɹ���');document.location.href('admin_culture_fl.asp');</script>"
end if
end if
%>
