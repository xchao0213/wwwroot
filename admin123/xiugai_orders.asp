<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%
if session("qx")=2 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
response.End
End If
%>
<% 
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
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>�鿴����</title>
</head>
<body>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�鿴����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<form name="add" method="post" action="?id=<%=rs("id")%>&xiugai=ok">
	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8" >
        <td height="28" width="16%" class="td">�������</td>
        <td width="84%"  class="td"><%=rs("OrderNo")%><input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" width="16%" class="td">��Ʒ����</td>
        <td class="td"><%
set rss=server.createobject("adodb.recordset")
exec="select * from [products] where id="&rs("cpid")&"" 
rss.open exec,conn,1,1 
response.Write("<a href=""../ShowProducts.asp?id="&rs("cpid")&""" target=""_blank"">"&rss("title")&"</a>")
rss.close
set rss=nothing
%></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" width="16%" class="td">��������</td>
        <td class="td"><%=rs("number")%></td>
      </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" class="td">��ϵ������</td>
        <td class="td"><%=rs("name")%></td>
      </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">��ϵ�绰</td>
        <td class="td"><%=rs("tel")%></td>
      </tr>
   <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">�ջ���ַ</td>
      <td class="td"><%=rs("address")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">����˵��</td>
      <td class="td"><%=rs("sm")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" class="td">�ύ����</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">����״̬</td>
      <td class="td"><%
state=rs("state")
if state=1 then
response.Write("<font color=#FF0000>�¶���,δ����!</font>")
else
response.Write("<font color=#000000>�Ѿ�����!</font>")
end if
%></td>
    </tr>
	<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" class="td">��ע��Ϣ</td>
      <td class="td">
          <textarea name="bz" cols="50" rows="3" ></textarea></td>
    </tr>
	<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
      <td height="25" class="td">����״̬</td>
      <td class="td"><select name="state">
            <option selected="selected" value="">��������</option>
            <option value="1" <%if rs("state")="1" then%>selected<%end if%>>���Ϊ�¶���</option>
            <option value="2" <%if rs("state")="2" then%>selected<%end if%>>���Ϊ�Ѵ���</option>
          </select></td>
    </tr>
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
      <td height="25" class="td">&nbsp;</td>
      <td class="td"><input type="button" name="Submit3" value="���ض����б�" onclick="window.location.href='admin_orders.asp' "  class="btn"/>  <input type="submit" name="button" id="button" value="������"  class="btn"/></td>
    </tr>
    </table>
	</form>
</td>
  </tr>
</table>
</body>
</html>
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
rs("bz")=bz
rs("state")=state
rs.update 
rs.close 
response.write "<script>alert('�����޸ĳɹ���');window.location.href='admin_orders.asp';</script>" 
end if
%> 