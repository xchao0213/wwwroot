<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<!--#include file="md5.Asp" -->
<%call chkAdmin(28)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>��Ա����</title>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��Ա����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" /></td>
<td width="5%" align="center" class="td">ID</td>
<td width="17%" height="25" align="center" class="td">��Ա�ʺ�</td>
<td width="16%" align="center" class="td">��Ա����</td>
<td width="14%" align="center" class="td">��½����</td>
<td width="14%" align="center" class="td">�Ƿ����</td>
<td width="17%" align="center" class="td">ע��ʱ��</td>
<td width="7%" align="center" class="td">�޸�</td>
<td width="7%" align="center" class="td">ɾ��</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?del=checkbox"> 
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from [user] order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">���޻�Ա!</div>")
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
for i=1 to rs.pagesize  %> 
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="5%" align="center" class="td"><input type="text" readonly="readonly" style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="17%" height="25" align="center" class="td"><a href="xiugai_user.asp?id=<%=rs("id")%>" style="color:#003399"><%=rs("useradmin")%></a> </td>
<td width="16%" align="center" class="td"><%=rs("zsname")%></td>
<td width="14%" align="center" class="td">��½������<%=rs("dlcs")%></td>
<td width="14%" align="center" class="td"><%if rs("sh")=1 then
response.Write("�����")
else
response.Write("<font color=#FF0000>δ���</font>")
end if%></td>
<td width="17%" align="center" class="td"><%=rs("data")%></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit3" value="�޸�" onclick="window.location.href='admin_user.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="3%" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="5%" align="center" class="td">ȫѡ</td>
<td height="25" colspan="7" class="td">
<input type="submit" class="btn" onclick="form.action='?sh=ok';" value="�������"/>
<input type="submit" class="btn" onclick="form.action='?sh=no';" value="ȡ�����"/>
<input type="submit" class="btn" onclick="form.action='?del=ok';" value="����ɾ��"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>
</td>
  </tr>
</table>
<% End If %>
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [user] where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<form  name="add" method="post" action="admin_user.asp?action=xiugai&id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�鿴/�޸Ļ�Ա����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="13%" height="28" align="right" class="td">�˺�ID</td>
        <td width="87%"  class="td"><%=rs("id")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr>
        <td width="13%" height="25" align="right" class="td">��Ա�ʺ�</td>
        <td class="td"><%=rs("useradmin")%></td>
      </tr>
      <tr>
        <td width="13%" height="25" align="right" class="td">��Ա����</td>
        <td class="td"><input name="userpassword" type="text" size="30"  /> ���޸�������! </td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">��ʵ����</td>
        <td class="td"><%=rs("zsname")%></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">��Ա�Ա�</td>
        <td class="td"><input type="radio" name="sex" value="0" <%if rs("sex")=0 then%>checked<%end if%>>������ 
<input type="radio" name="sex" value="1" <%if rs("sex")=1 then%>checked<%end if%>>Ůʿ</td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">���뱣������</td>
        <td class="td"><%=rs("wen")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">���뱣����</td>
      <td class="td"><%=rs("da")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��˾����</td>
      <td class="td"><label>
        <input name="gsname" type="text" value="<%=rs("gsname")%>" size="40" />
      </label></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��˾��ַ</td>
      <td class="td"><input name="gsadd" type="text" value="<%=rs("gsadd")%>" size="40" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��������</td>
      <td class="td"><input name="youbian" type="text" value="<%=rs("youbian")%>" size="20" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��ϵ�绰</td>
      <td class="td"><input name="tel" type="text" value="<%=rs("tel")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��ϵ����</td>
      <td class="td"><input name="fax" type="text" value="<%=rs("fax")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��ϵ�ֻ�</td>
      <td class="td"><input name="sj" type="text" value="<%=rs("sj")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��������</td>
      <td class="td"><input name="mail" type="text" value="<%=rs("mail")%>" size="30" /></td>
    </tr>
  <tr>
      <td height="25" align="right" class="td">��˾��ַ</td>
      <td class="td"><input name="wz" type="text" value="<%=rs("wz")%>" size="30" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">ע��ʱ��</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
  <tr>
      <td height="25" align="right" class="td">����½</td>
      <td class="td"><%=rs("dldata")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��½����</td>
      <td class="td"><%=rs("dlcs")%> ��</td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">ע��IP</td>
      <td class="td"><%=rs("ip")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">&nbsp;</td>
      <td class="td"><input type="submit" name="button" id="button" value="��������" class="btn"/></td>
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
sql="select * from [user] where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('ɾ���ɹ���');window.location.href='admin_user.asp';</script>"
end if 
%>
<%
if Request.QueryString("sh")="ok" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');window.location.href='admin_user.asp';</script>" 
response.end()
end if
sql="update [user] set sh=1 where id in ("&Request("id")&")" 
conn.execute(sql)
Response.Write "<script>alert('��ϲ!��˳ɹ�!');window.location.href='admin_user.asp';</script>" 
end if
 
if Request.QueryString("sh")="no" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');window.location.href='admin_user.asp';</script>" 
response.end()
end if
sql="update [user] set sh=0 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ȡ����˳ɹ�!');window.location.href='admin_user.asp';</script>" 
end if

if Request.QueryString("del")="ok" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');window.location.href='admin_user.asp';</script>" 
response.end()
end if
sql="delete from [user] where id in ("&Request("id")&")"
conn.Execute ( sql )
Response.Write "<script>alert('��ϲ!�����ɹ�!');window.location.href='admin_user.asp';</script>" 
end if

'=====================
if Request.QueryString("xiugai")="ok" then 
id=request("id")
userpassword=request.form("userpassword")
sex=request.form("sex")
gsname=request.form("gsname")
gsadd=request.form("gsadd")
youbian=request.form("youbian")
tel=request.form("tel")
fax=request.form("fax")
sj=request.form("sj")
mail=request.form("mail")
wz=request.form("wz")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from [user] where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
if request.form("userpassword") <> "" then
rs("userpassword")=md5(request.form("userpassword"))
end if
rs("sex")=sex
rs("gsname")=gsname
rs("gsadd")=gsadd
rs("youbian")=youbian
rs("tel")=tel
rs("fax")=fax
rs("sj")=sj
rs("mail")=mail
rs("wz")=wz
rs.update 
rs.close 
response.write "<script>alert('�����޸ĳɹ���');window.location.href='admin_user.asp?action=admin';</script>" 
end if%>