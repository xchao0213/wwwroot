<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>����Ա����</title>
<SCRIPT language=javascript>
function SelectAll(form){
  for (var i=0;i<form.AdminMight.length;i++){
    var e = form.AdminMight[i];
    if (e.disabled==false)
       e.checked = form.chkAll.checked;
    }
}
</script>
</head>
<body>
<%if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">����Ա����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="8%" align="center" class="td">ID</td>
<td width="16%" height="25" align="center" class="td">����Ա</td>
<td width="18%" align="center" class="td">Ȩ ��</td>
<td width="15%" align="center" class="td">��½����</td>
<td width="26%" align="center" class="td">����½</td>
<td width="7%" align="center" class="td">�޸� </td>
<td width="10%" align="center" class="td">ɾ��</td>
</tr></thead>
<%	
set rs=server.createobject("adodb.recordset")
if key=0 then
exec="select * from admin order by id asc" 
else
exec="select * from admin where id="&session("admin")&"" 
end if
rs.open exec,conn,1,1 
if rs.eof then
response.Write "������Ϣ��"
else
rs.PageSize =10 'ÿҳ��¼����
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
for i=1 to rs.pagesize
if rs("key")<>0 then
key="<font color=#336699>��ͨ����Ա</font>"
else
key="<font color=#336699>��������Ա</font>"
end if%>
<tr>
<td width="8%" align="center" class="td"><%=rs("id")%></td>
<td width="16%" height="25" align="center" class="td"><%=rs("admin")%> </td>
<td width="18%" align="center" class="td"><%=key%></td>
<td width="15%" align="center" class="td"><%=rs("dlcs")%></td>
<td width="26%" align="center" class="td"><%=rs("dldata")%></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit3" value="�޸�" onClick="window.location.href='admin_administrator.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/> </td>
<td width="10%" align="center" class="td"><input type="button" name="Submit" value="ɾ��" onClick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='admin_administrator.asp?action=admin&id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<% rs.movenext 
if rs.eof then exit for 
next 
end if%></table>
</td>
  </tr>
</table>
<div style="margin-top:10px">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?add=ok" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">���ӹ���Ա</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="12%" height="25" align="right" class="td">����Ա�ʺ�</td>
        <td width="88%"  class="td"><input name="admin" type="text" size="30"  /> * (����ʹ������!)</td>
      </tr>
      <tr>
        <td width="12%" height="13" align="right" class="td">��½����</td>
        <td class="td"><input name="password" type="text" size="30"  />
          * </td>
      </tr>
      <tr>
        <td width="12%" height="12" align="right" class="td">ȷ������</td>
        <td class="td"><input name="password3" type="text" size="30"  />
          * </td>
      </tr>
      <tr>
        <td width="12%" height="25" align="right" class="td">��ʵ����</td>
        <td class="td"><input name="zsname" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">����ԱQQ</td>
        <td class="td"><input name="qq" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">����Ա����</td>
        <td class="td"><label>
          <select name="key" id="key">
            <option value="1">��ͨ����Ա</option>
            <option value="0">��������Ա</option>
          </select></label></td>
      </tr>
      <tr>
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="�ύ����"  class="btn"/></td>
      </tr>
    </table></td>
  </tr></form>
</table>
</div>
<%end if%>
<!--����Ϊ��������Ա�޸�����-->
<%if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from admin where id="&id
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 %> 
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="admin_administrator.asp?action=xiugai<%=rs("id")%>&adminxiugai=ok" method="post" name="add">
   <tr>
     <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸Ĺ���Ա����</div></td>
   </tr>
   <tr>
     <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
         <tr >
           <td width="12%" height="25" align="right" class="td">����Ա�ʺ�</td>
           <td width="88%"  class="td"><input name="id" type="hidden" value="<%=rs("id")%>" size="30"  />
		   <input name="admin" type="text" value="<%=rs("admin")%>" size="30"  /></td>
        </tr>
         <tr>
           <td width="12%" align="right" class="td">��½����</td>
           <td class="td"><input name="password" type="text" size="30"  /> ���޸������գ�</td>
         </tr>
            <tr>
           <td width="12%" align="right" class="td">ȷ������</td>
           <td class="td"><input name="password2" type="text" size="30"  /></td>
         </tr>
        <tr>
           <td width="12%" height="25" align="right" class="td">��ʵ����</td>
           <td class="td"><input name="zsname" type="text" value="<%=rs("zsname")%>" size="30"  /></td>
         </tr>
              <tr>
           <td height="25" align="right" class="td">����ԱQQ</td>
           <td class="td"><input name="qq" type="text" value="<%=rs("qq")%>" size="30"  /></td>
         </tr>
         <tr>
           <td height="25" align="right" class="td">����Ա����</td>
           <td class="td">
<%if key=0 then%><input type="radio" name="key" value="0" <%if rs("key")=0 then%>checked<%end if%> onClick="Date3.style.display='none'">��������Ա<%end if%> 
<input type="radio" name="key" value="1" <%if rs("key")=1 then%>checked<%end if%> onClick="Date3.style.display=''">��ͨ����Ա�� 
 </td>
         </tr>
<%'if key=0 then%>
<tr id="Date3" style="display:<%If rs("key")=0 Then Response.Write "none"%>">
<td height="25" colspan="2"  class="td">
<% dim MightList(5,10)
MightList(0,0) = "��վ����"
MightList(0,1) = "ȫ�ֲ���-1"
MightList(0,2) = "���Ӳ���-2"
MightList(0,3) = "�����˵�-3" 
MightList(0,4) = "����Ա����-4"
MightList(0,5) = "��½��¼-5"
MightList(0,6) = "SQL����-6"
MightList(0,7) = "���ɾ�̬-7"
MightList(0,8) = "�ռ�ռ�ò鿴-8"

MightList(1,0)="���ݹ���1"
MightList(1,1)="��ҳ����-9"
MightList(1,2)="���Ź���-10"
MightList(1,3)="���ŷ���-10"
MightList(1,4)="��������-12"
MightList(1,5)="��������-13"
MightList(1,6)="��Ʒ����-14"
MightList(1,7)="��Ʒ����-15"
MightList(1,8)="�Ŷӹ���-16"
MightList(1,9)="�������-17"
MightList(1,10)="��Ƶ����-18"

MightList(2,0)="���ݹ���2"
MightList(2,1)="��Ƶ����-19"
MightList(2,2)="���ع���-20"
MightList(2,3)="���ط���-21"
MightList(2,4)="������-22"
MightList(2,5)="������-23"
MightList(2,6)="�˲Ź���-24"

MightList(3,0)="��������"
MightList(3,1)="JS���-25"
MightList(3,2)="��ҳ��ͼ-26"
MightList(3,3)="Ƶ���õ�-27"
MightList(3,4)="��Ա����-28"
MightList(3,5)="���߿ͷ�-29"
MightList(3,6)="��������-30"
MightList(3,7)="���ݱ���-31"
MightList(3,8)="��������-32"

MightList(4,0)="��������"
MightList(4,1)="���Թ���-33"
MightList(4,2)="��ͼ����-34"
MightList(4,3)="��������-35"
MightList(4,4)="��վ�ƹ�-36"
MightList(4,5)="SEO����-37"%>   
<fieldset style="margin:2px 2px 2px 2px">
<legend><B>ϵͳȨ������</B>	<label><input name="chkAll" type="checkbox" id="chkAll" onClick="SelectAll(this.form)">ȫѡ</label></legend>
<%
Dim n,i
for i=0 to ubound(MightList,1)
If i<>0 then
Response.Write "<br><br>"
End if
Response.Write "<b>"&MightList(i,0)&"</b><br>"
for	j=1	to ubound(MightList,2)
if isempty(MightList(i,j)) then exit for
tmpmenu=split(MightList(i,j),"-")
menuname=tmpmenu(0)
menurl=tmpmenu(1)%>
    <label class="box"><input name="manage" type="checkbox" id="AdminMight" value="<%=menurl%>"<% if instr(","& rs("manage") &",",","&menurl&",")>0 then response.write " checked" %>><%=menurl%>.<%=menuname%></label>
    <%next
next%>
</fieldset></td>
</tr>
<%' End If %>
         <tr>
           <td height="25" align="right" class="td">&nbsp;</td>
           <td class="td"><input type="submit" name="button2" id="button2" value="ȷ���޸�"  class="btn"/></td>
         </tr>
     </table></td>
   </tr></form>
 </table>
<%end if%>
</body>
</html>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,3
admin=request.form("admin")
password=request.form("password")
password3=request.form("password3")
zsname=request.form("zsname")
qq=request.form("qq")
key=request.form("key")
if admin=""  then 
response.Write("<script language=javascript>alert('��½�ʺŲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
letters="0123456789abcdefghijklmnopqrstuvwxyz" 
admin=Lcase(trim(Request.Form("admin"))) 
for i=1 to len(admin) 
u=mid(admin,i,1) 
if Instr(letters,u)=0 then 
response.write "<script>alert('��½�ʺ�ֻ������ĸ�����ּ��»������!');history.go(-1);</script>" 
response.end 
end if 
next 
if password=""  then 
response.Write("<script language=javascript>alert('��½���벻��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
if request.Form("password")<>request.Form("password3") then 
Response.Write "<script>alert('�������벻һ�£����������룡');history.go(-1);</script>" 
response.end
end if
rs("admin")=admin
rs("password")=md5(request.form("password"))
rs("zsname")=zsname
rs("qq")=qq
rs("key")=key
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('��ӳɹ���');window.location.href='admin_administrator.asp?action=admin';</script>" 
end if

if request("adminxiugai")="ok" then
id=request("id")
if id="" or not isnumeric(id) then
Response.Write "<script language='javascript'>alert('��������!');document.location.href('admin_administrator.asp?action=admin');</script>"
Response.End()
end if
SQL="Select * from admin where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script language='javascript'>alert('��������ȷ!');document.location.href('admin_administrator.asp?action=admin');</script>"
Response.End()
end if
if request.form("password")<>"" then
rs("password")=md5(request.form("password"))
else
end if
if request.Form("password")<>request.Form("password2") then 
Response.Write "<script>alert('�������벻һ�£����������룡');history.go(-1);</script>" 
response.end
end if
rs("admin")=request.form("admin")
rs("zsname")=request.form("zsname")
rs("qq")=request.form("qq")
if key=0 then
rs("manage")=replace(request.form("manage")," ","")
rs("key")=request.form("key")
end if
rs.update 
rs.close 
response.write "<script>alert('����Ա�����޸ĳɹ���');window.location.href='admin_administrator.asp?action=admin';</script>" 
end if

if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from admin where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('ɾ���ɹ���');window.location.href='admin_administrator.asp?action=admin';</script>"
end if %>