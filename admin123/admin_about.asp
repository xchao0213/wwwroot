<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(9)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>��ҳ����</title>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��ҳ����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" /></td>
<td width="5%" align="center" class="td">ID</td>
<td width="13%" height="25" align="center" class="td">����</td>
<td width="48%" align="center" class="td">��ַ</td>
<td width="15%" align="center" class="td">����</td>
<td width="7%" align="center" class="td">�޸�</td>
<td width="9%" align="center" class="td">ɾ��</td>
</tr></thead>
<form action="?" method="post" name="form">
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from about order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">���޼�¼!</div>")
else
rs.PageSize =30 'ÿҳ��¼����
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
<td width="13%" height="25" align="center" class="td"><a href="xiugai_about.asp?id=<%=rs("id")%>" style="color:#003399"><%=rs("title")%></a> </td>
<td width="48%" align="center" class="td">��ַ��
<input value="/about/?id=<%=rs("id")%>" /> 
[<a href="/about/?id=<%=rs("id")%>" target="_blank" style="color:#003399">����</a>]</td>
<td width="15%" align="center" class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" style="text-align:center; width:40px"/></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit3" value="�޸�" onclick="window.location.href='xiugai_about.asp?id=<%=rs("id")%>' "  class="btn"/></td>
<td width="9%" align="center" class="td">
<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='admin_about.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<% 
rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr><td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
  <td align="center">ȫѡ</td>
  <td colspan="5"><input type="submit" class="btn" onclick="form.action='?action=del';" value="ɾ��"/>
  <input type="submit" class="btn" onclick="form.action='?action=ordsc';" value="��������"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing%></td>
  </tr>
</form>
</table>
</td>
</tr>
</table>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from about where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('ɾ���ɹ���');window.location.href='admin_about.asp';</script>"
end if 
'==��������
if Request.QueryString("action")="ordsc" then'����
id=Split(request.form("id"),", ")
px_id=Split(request.form("px_id"),", ")
set rs=server.createobject("adodb.recordset")
sql="select * from about order by px_id asc"
rs.open sql,conn,1,3
for i=0 to Ubound(px_id)
rs("px_id")=px_id(i)
rs.update
rs.movenext
next
rs.close
set rs=nothing
Response.Write "<script>alert('��������ɹ�!');window.location.href='admin_about.asp';</script>" 
end if
'==����ɾ��
if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [about] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_about.asp';</script>" 
end if
%>
