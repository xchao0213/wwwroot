<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(30)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>��������</title>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#cccccc">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
      <form id="form1" name="form1" method="post" action="admin_orders.asp?del=checkbox">
<div style="text-align:left;height:25px; padding:0px 10px 0px 10px;"> 
      <select name="ssfl" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
      <option>---����鿴---</option>
      <option value="?action=admin&lxid=">---ȫ������---</option>
      <option value="?action=admin&lxid=1">---�¶���---</option>
      <option value="?action=admin&lxid=2">---�Ѵ���---</option>
      </select></div>
<thead><tr>
  <td width="48" align="center" class="td"><input name="ID" type="checkbox"/></td>
  <td width="66" align="center" class="td">ID</td>
  <td width="397" align="center" class="td">������</td>
  <td width="199" align="center" class="td">�������</td>
  <td width="291" height="25" align="center" class="td">����״̬</td>
<td width="344" align="center" class="td">����ʱ��</td>
<td width="129" align="center" class="td">������</td>
<td width="120" align="center" class="td">ɾ������</td>
</tr></thead>
<%id=request.QueryString("lxid") 
set rs=server.createobject("adodb.recordset")
if id<>"" then
exec="select * from [orders] where state="&id&" order by id desc" 
else
exec="select * from [orders] order by id desc"  
end if
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">���޼�¼!</div>")
else
rs.PageSize =20 'ÿҳ��¼����
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
  <td width="48" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
  <td width="66" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
  <td width="397" align="center" class="td"><a href="?action=xiugai&id=<%=rs("id")%>"><b><%=rs("OrderNo")%></b></a></td>
  <td width="199" align="center" class="td"><%=rs("rmb")%></td>
  <%if rs("state")=1 then
  ddzt="�ȴ���Ҹ���"
  elseif rs("state")=2 then
  ddzt="����Ѹ���ȴ�����"
  elseif rs("state")=3 then
  ddzt="�����ѷ������ȴ�ȷ���ջ�"
  elseif rs("state")=8 then
  ddzt="�������"
  else
  ddzt="���������쳣"
  end if %> 
  <td width="291" height="25" align="center" class="td"><%=ddzt%></td>
<td width="344" align="center" class="td"><%=rs("data")%></td>
<td width="129" align="center" class="td">
<input type="button" name="Submit3" value="������" onclick="window.location.href='?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="120" align="center" class="td">
<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ���ö�����ɾ���󲻿ɻָ�!')){window.location.href='admin_orders.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/>
</td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="48" height="30" align="center" style="padding-left:5px;">
<input type="checkbox" name="allbox" onclick="CheckAll()" />
</td>
<td width="66" height="30" align="center" style="padding-left:5px;">ȫѡ</td>
<td colspan="2"><label>
  <select name="lxid">
  <option selected="selected" value="">��������</option>
  <option value="1">���Ϊδ����</option>
  <option value="2">���Ϊ�Ѹ���</option>
  <option value="3">���Ϊ�ѷ���</option>
  <option value="8">���Ϊ�꽻��</option>
  <option value="4">�����쳣</option>
  </select>
  <input type="submit" name="button" id="button" value="�ύ"  class="btn"/>
  </label>
  <%heji=conn.execute("Select Sum(rmb) as rmb From orders where state=8")(0)+conn.execute("Select Sum(kuaidi) as kuaidi From orders where state=8")(0)%>
  �ۼ�����:<font style="font-size:24px; color:#F00; font-family:'Trebuchet MS', Arial, Helvetica, sans-serif">��<%=heji%></font> Ԫ
  
</td>
<td colspan="4"><%call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%></td>
</form>  
    </table></td>
  </tr>
</table>
<% end if %>
<% if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [orders] where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#ccccccc">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�鿴����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<form name="add" method="post" action="?id=<%=rs("id")%>&xiugai=ok">
	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="stable" >
      <tr >
        <td width="10%" height="28" align="right" class="td">�������</td>
        <td width="90%" class="td"><%=rs("OrderNo")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">��Ʒ����</td> 
        <td class="td"><%
set rss=server.createobject("adodb.recordset")
exec="select * from [products] where id in ("&rs("cpid")&")" 
rss.open exec,conn,1,1 
if rss.eof then
response.write ("<a style=""color:#FF0000"">�޴˲�Ʒ��˲�Ʒ��ɾ��!</a>")
else
while not rss.eof
response.Write("<a href=""../Products/ShowProducts.asp?id="&rs("cpid")&""" target=""_blank"">"&rss("title")&"</a><br>")
rss.movenext
wend
rss.close
end if
set rss=nothing
%></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">��������</td>
        <td class="td"><%=rs("number")%></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">������</td>
        <td class="td"><input name="rmb" type="text" value="<%=rs("rmb")%>"<% If rs("state")<>1 Then %> style=" background:#E4E4E4" readonly<% End If %> /></td>
      </tr>
      <tr>
        <td width="10%" height="25" align="right" class="td">��ݷ���</td>
        <td class="td"><input name="kuaidi" type="text" value="<%=rs("kuaidi")%>"<% If rs("state")<>1 Then %> style=" background:#E4E4E4" readonly<% End If %> /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">��ϵ������</td>
        <td class="td"><%=rs("name")%></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">��ϵ�绰</td>
        <td class="td"><%=rs("tel")%></td>
      </tr>
   <tr>
      <td height="25" align="right" class="td">�ջ���ַ</td>
      <td class="td"><%=rs("address")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">����˵��</td>
      <td class="td"><%=rs("sm")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">�ύ����</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">����״̬</td>
      <td class="td"><%
state=rs("state")
if state=1 then
response.Write("<font color=#FF0000>�¶���,δ����!</font>")
else
response.Write("<font color=#000000>�Ѿ�����!</font>")
end if
%></td>
    </tr>
	<tr>
      <td height="25" align="right" class="td">��ע��Ϣ</td>
      <td class="td"><textarea name="bz" cols="50" rows="3" ></textarea></td>
    </tr>
	<tr>
      <td height="25" align="right" class="td">��������</td>
      <td class="td"><select name="kddm">
            <option selected="selected" value="">��ݹ�˾</option>
			<%set rsc=server.CreateObject("adodb.recordset")
            rsc.open "select * from orders_kd order by px_id asc",conn,1,1
            while not rsc.eof
            if rs("kddm")=rsc("kddm") then
            response.Write("<option value="""&rsc("kddm")&""" selected>" & rsc("title") & "</option>")
            else
            response.Write("<option value="""&rsc("kddm")&""">" & rsc("title") & "</option>")
            end if
            rsc.movenext
            wend
            rsc.close
            set rsc=nothing%>
          </select><input name="kddh" type="text" value="<%=rs("kddh")%>" placeholder="�������ݵ���" /></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">����״̬</td>
      <td class="td"><select name="state">
            <option selected="selected" value="">��������</option>
            <option value="1" <%if rs("state")="1" then%>selected<%end if%>>���Ϊδ����</option>
            <option value="2" <%if rs("state")="2" then%>selected<%end if%>>���Ϊ�Ѹ���</option>
            <option value="3" <%if rs("state")="3" then%>selected<%end if%>>���Ϊ�ѷ���</option>
            <option value="8" <%if rs("state")="8" then%>selected<%end if%>>���Ϊ�꽻��</option>
            <option value="4" <%if rs("state")="4" then%>selected<%end if%>>�����쳣</option>
          </select></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">&nbsp;</td>
      <td class="td"><input type="button" name="Submit3" value="���ض����б�" onclick="window.location.href='admin_orders.asp' "  class="btn"/>  <input type="submit" name="button" id="button" value="������"  class="btn"/></td>
    </tr>
    </table>
	</form>
</td>
  </tr>
</table>
<% end if %>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from orders where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰ����ɾ���ɹ���');window.location.href='admin_orders.asp?action=admin';</script>"
end if 
%>
<%
if Request.QueryString("del")="checkbox" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');window.location.href='admin_orders.asp?action=admin';</script>" 
response.end()
end if
dim sql
lx=request.Form("lxid")
if lx="" then
Response.Write "<script>alert('����!��ѡ���������!');window.location.href='admin_orders.asp?action=admin';</script>" 
response.end
end if
if lx=1 then
sql="update orders set state=1 where id in ("&Request("id")&")" 
conn.execute(sql) 
elseif lx=2 then 
sql="update orders set state=2 where id in ("&Request("id")&")" 
conn.execute(sql)
elseif lx=3 then
sql="update orders set state=3 where id in ("&Request("id")&")" 
conn.execute(sql) 
elseif lx=4 then 
sql="update orders set state=4 where id in ("&Request("id")&")" 
conn.execute(sql) 
else
sql="delete from orders where id in ("&Request("id")&")"
conn.Execute ( sql )
end if
conn.close
set conn=nothing
Response.Write "<script>alert('��ϲ!�����ɹ�!');window.location.href='admin_orders.asp?action=admin';</script>" 
end if
%>
<%
if Request.QueryString("xiugai")="ok" then 
id=request("id")
bz=request.form("bz")
state=request.form("state")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from orders where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
rs("rmb")=request.form("rmb")
rs("kuaidi")=request.form("kuaidi")
rs("kddm")=request.form("kddm")
rs("kddh")=request.form("kddh")
rs("bz")=bz
rs("state")=state
rs.update 
rs.close 
response.write "<script>alert('�����޸ĳɹ���');window.location.href='admin_orders.asp?action=admin';</script>" 
end if
%> 