<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(24)%>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
    
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<form id="form1" name="form1" method="post" action="?"> 
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox"/></td>
<td width="4%" align="center" class="td">ID</td>
<td width="17%" height="25" align="center" class="td">����</td>
<td width="16%" align="center" class="td">ְλ</td>
<td width="17%" align="center" class="td">�� ��</td>
<td width="11%" align="center" class="td">ѧ��</td>
<td width="18%" align="center" class="td">ʱ��</td>
<td width="7%" align="center" class="td">�鿴</td>
<td width="7%" align="center" class="td">ɾ��</td>
</tr></thead>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from Resume order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">��û���յ��κμ���!</div>")
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
for i=1 to rs.pagesize%>
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="4%" align="center" class="td"><input type="text" readonly style="text-align:center;width:40px" value="<%=rs("id")%>"/></td>
<td width="17%" height="25" align="center" class="td"><a href="admin_Resume.asp?action=ck&id=<%=rs("id")%>" style="color:#003399"><%=rs("name")%></a></td>
<td width="16%" align="center" class="td"><%=rs("ypzw")%></td>
<td width="17%" align="center" class="td"><%=rs("sex")%></td>
<td width="11%" align="center" class="td"><%=rs("xueli")%></td>
<td width="18%" align="center" class="td"><%=rs("data")%></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit3" value="�鿴" onclick="window.location.href='admin_Resume.asp?action=ck&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="7%" align="center" class="td">

<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">ȫѡ</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?action=del';" value="ɾ��"/><%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>        
   </td>
  </tr>
</table>
<%end if%>
<%if request.querystring("action")="ck" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from Resume where id="&id
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�鿴����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="11%" height="28" align="right" class="td">ӦƸְλ��</td>
        <td width="89%"  class="td"><%=rs("ypzw")%></td>
      </tr>
      <tr>
        <td width="11%" height="25" align="right" class="td">������</td>
        <td class="td"><%=rs("name")%></td>
      </tr>
      <tr>
        <td width="11%" height="25" align="right" class="td">�Ա�</td>
        <td class="td"><%=rs("sex")%></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">���䣺</td>
        <td class="td"><%=rs("nn")%></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">���壺</td>
        <td class="td"><%=rs("mz")%></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">������</td>
        <td class="td"><%=rs("hj")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">����״����</td>
      <td class="td"><%=rs("hyzk")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">��ߣ�</td>
      <td class="td"><%=rs("sg")%></td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">���أ�</td>
      <td class="td"><%=rs("tz")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">���֤��</td>
      <td class="td"><%=rs("sfz")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">ѧ����</td>
      <td class="td"><%=rs("xueli")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">�����ڵأ�</td>
      <td class="td"><%=rs("szd")%></td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">��ҵԺУ��</td>
      <td class="td"><%=rs("byyx")%></td>
    </tr>
   <tr>
     <td height="25" align="right" class="td">��ϵ�绰��</td>
     <td class="td"><%=rs("tel")%></td>
   </tr>
   <tr>
     <td height="25" align="right" class="td">��ϵ�ֻ���</td>
     <td class="td"><%=rs("sj")%></td>
   </tr>
    <tr>
      <td height="25" align="right" class="td">����������</td>
      <td class="td"><%=rs("jybj")%></td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">�������飺</td>
      <td class="td"><%=rs("gzjn")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">ר����</td>
      <td class="td"><%=rs("zc")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">�ύIP��</td>
      <td class="td"><%=rs("ip")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">�ύ���ڣ�</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
   <tr>
      <td height="25" class="td"></td>
      <td class="td"><input type="button" name="Submit3" value="�����б�" onclick="window.location.href='admin_Resume.asp?action=admin' "  class="btn"/></td>
    </tr>
    
    </table></td>
  </tr>
</table>
<%end if%>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from Resume where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('ɾ���ɹ���');window.location.href='admin_Resume.asp';</script>"
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [case] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_Case.asp';</script>" 
end if

%>
