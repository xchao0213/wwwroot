<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(10)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>԰������</title>
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
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">԰������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<div style="text-align:left; height:25px; padding:2px 10px 2px 10px;"> 
<%set rsk=server.CreateObject("adodb.recordset")
rsk.open "select * from fengfa_fl order by id asc",conn,1,1
if rsk.eof then
response.write""
else
Response.Write("<select name=""ssfl"" id=""jumpMenu"" onChange=""window.open(this.options[this.selectedIndex].value,'_self')"">")
Response.Write("<option>---����鿴---</option>")
response.Write("<option value=""?id="">ȫ��԰��</option>")
while not rsk.eof
response.Write("<option value=""?id=" & rsk("id") & """>" & rsk("title") & "</option>")
rsk.movenext
wend
Response.Write("</select>")
end if
rsk.close
set rsc=nothing%>
</div>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="50" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="50" align="center" class="td">����ID</td>
<td width="56" align="center" class="td">��������</td>
<td width="70" height="25" align="center" class="td">����ͼ</td>
<td height="25" align="center" class="td">�Ļ�����</td>
<td width="60" align="center" class="td">����</td>
<td width="60" align="center" class="td">�Ƽ�</td>
<td width="60" align="center" class="td">Ȩ��</td>
<td width="120" align="center" class="td">����ʱ��</td>
<td width="60" align="center" class="td">�޸�</td>
<td width="60" align="center" class="td">ɾ��</td>
</tr></thead>
<form id="form" name="form" method="post" action="?"> 
<%id= request.QueryString("id") 
set rs=server.createobject("adodb.recordset")
if id<>"" then
exec="select * from fengfa where ssfl="& id &"  order by id desc" 
else
exec="select * from fengfa order by id desc"  
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
for i=1 to rs.pagesize
if IsNull(rs("img")) or rs("img")="" then
img="<img src=""images/nopic.gif"" width=""50"" height=""25"" alt=""ͼƬ����ͼ""/>"
else
img="<img src="&rs("img")&" width=""50"" height=""25"" alt=""ͼƬ����ͼ""/>"
end if
if rs("tuijian")=1 then
tuijian="<font color=#FF0000>[�Ƽ�]</font>"
else
tuijian="[����]"
end if
if rs("qx")=9 then
qx="<font color=#FF0000>[��Ա]</font>"
else
qx="[�ο�]"
end if
%>
<tr>
<td align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td align="center" class="td">
<%set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from fengfa_fl",conn,1,1
while not rsc.eof
if rs("ssfl")=rsc("id") then
response.Write("<a href=""?id=" & rsc("id") & """><font color=#003399>[" & rsc("title") & "]</font></a>")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
%></td>
<td height="25" align="center" class="td"><%=img%></td>
<td height="25" class="td"><a href="xiugai_fengfa.asp?id=<%=rs("id")%>"><%=stvalue(DelHtml(rs("title")),40)%></a>  </td>
<td align="center" class="td"><%=rs("zz")%></td>
<td align="center" class="td"><%=tuijian%></td>
<td align="center" class="td"><%=qx%></td>
<td align="center" class="td"><%=rs("data")%></td>
<td align="center" class="td"><input type="button" name="Submit3" value="�޸�" onclick="window.location.href='xiugai_fengfa.asp?id=<%=rs("id")%>'"  class="btn"/></td>
<td align="center" class="td"><input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='admin_fengfa.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td height="30" align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" />
</td>
<td height="30" align="center">ȫѡ</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?page=<%=page%>&tuijian=ok';" value="�Ƽ�"/>
<input type="submit" class="btn" onclick="form.action='?page=<%=page%>&tuijian=no';" value="����"/>
<input type="submit" class="btn" onclick="form.action='?page=<%=page%>&action=del';" value="ɾ��"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
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
sql="select * from fengfa where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰ԰��ɾ���ɹ���');window.location.href='admin_fengfa.asp?page="&page&"';</script>"
end if 


if Request.QueryString("tuijian")="ok" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="update fengfa set tuijian=1 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!�Ƽ��ɹ�!');window.location.href='admin_fengfa.asp?page="&page&"';</script>" 
end if

if Request.QueryString("tuijian")="no" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="update fengfa set tuijian=0 where id in ("&Request("id")&")" 
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ȡ���Ƽ��ɹ�!');window.location.href='admin_fengfa.asp?page="&page&"';</script>" 
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from fengfa where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_fengfa.asp?page="&page&"';</script>" 
end if

conn.close
set conn=nothing%>
