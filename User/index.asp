<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!--#include file="../Include/seeion.asp" -->
<%if Request.ServerVariables("QUERY_STRING")="" then response.Redirect "?action=admin"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ա����_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>��Ա���� <i>User</i></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>User/">��Ա����</a>��Ա��ҳ</span></div>
<div class="w_650 fl">
<div class="position">
<%if request.querystring("action")="admin" then%>
<div  class="ubox">
<table width="100%" border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse:collapse;line-height:25px; text-indent:3px">
<tr>
<td width="52%" colspan="2">
<% 
set rs=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rs.open exec,conn,1,1 
response.write(""&rs("zsname")&"<a href=""../User/index.asp"">["&rs("useradmin")&"]</a>&nbsp;��ӭ����")
%></td>
    </tr>
  <tr>
    <td width="52%" align="center"><a href="../User/?action=Edit">�ҵ�����</a></td>
    <td width="48%" align="center"><a href="../User/my_orders.asp">�ҵĶ���</a></td>
  </tr>
  <tr>
    <td align="center"><a href="../User/?action=password">�޸�����</a></td>
    <td align="center"><a href="../User/loginOUT.asp">�˳���½</a></td>
  </tr>
</table>
</div>
<div class="ubox">
<div class="pronav"><h1>�������</h1></div>
<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse: collapse;line-height:25px; text-indent:3px">
<tr>
    <td width="19%" align="center"><strong>�������</strong></td>
    <td width="27%" align="center"><strong>��Ʒ����</strong></td>
    <td width="21%" align="center"><strong>ʱ ��</strong></td>
    <td width="21%" align="center"><strong>״ ̬</strong></td>
    <td colspan="2" align="center"><strong>�� ��</strong></td>
    </tr>
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from [orders] where userid="&session("username")&" order by id desc "  
rs.open exec,conn,1,1 
if rs.eof then
response.Write "<tr><td colspan=""5"">���޶�����</td></tr>"
end if
while not rs.eof
if rs("state")=1 then
ddzt="<a href="""&dir&"alipay/"&alipayLX&"/alipayapi_re.asp?id="&rs("id")&""" title=""���������ɸ���""  target=""_blank"" class=""kfal_a"" onClick=""document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'""><font>��������</font></a>"
elseif rs("state")=2 then
ddzt="����Ѹ���ȴ�����"
elseif rs("state")=3 then
ddzt="�ѷ������ȴ�ȷ���ջ�"
elseif rs("state")=8 then
ddzt="�������"
else
ddzt="���������쳣"
end if%>
  <tr>
    <td width="19%" align="center"><a href="my_Showorders.asp?id=<%=rs("id")%>" title="����鿴��������"><%=rs("OrderNo")%></a></td>
    <td width="27%"><%=stvalue(rs("title"),24)%></td>
    <td width="21%" align="center"><%=rs("data")%></td>
    <td width="21%" align="center"><%=ddzt%></td>
    <td width="6%" align="center"><a href="my_Showorders.asp?id=<%=rs("id")%>">�鿴</a></td>
    <td width="6%" align="center"><a href="javascript:" onClick="javascript:if(confirm('ȷ��ɾ���ö�����ɾ���󲻿ɻָ�!')){window.location.href='?action=admin&id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}" title="ɾ���˶���">ɾ��</a></td>
  </tr>
<%rs.movenext
wend
rs.close
set rs=nothing%>
</table>
</div>

<div class="ubox">
<div class="pronav"><h1>��Ʒ�Ƽ�</h1></div>
<%set rs=server.createobject("adodb.recordset") 
exec="select top 3  * from products where tuijian=1 order by px_id desc,id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">���޲�Ʒ</div>"
else
while not rs.eof
if html=0 then
url=""&Dir&"Products/"&ShowProducts&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Products/ShowProducts.asp?id="&rs("id")&""
end if%>
<div class="Pro_box">
<a href="<%=url%>" title="<%=rs("title")%>"><img src="<%=rs("img")%>" width="234" height="150" alt="<%= rs("title") %>" onerror="this.src='images/nologo.jpg'"/><p><%= stvalue(rs("title"),21) %></p></a>
</div>
<%rs.movenext
wend
end if
rs.close
set rs=nothing%>
</div>
<%end if
if request.querystring("action")="Edit" then%>
<div  class="ubox">
<div class="pronav"><h1>�ҵ�����</h1></div>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rs.open exec,conn,1,1 %>
<form action="?action=Edit&id=<%=rs("id")%>&xiugai=info" method="post" name="add">
<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse:collapse; margin-top:10px; text-indent:3px">  
	<tr>
      <td width="18%" align="right">��½�ʺţ�</td>
      <td height="30"><%=rs("useradmin")%>
        <input name="id" type="hidden" value="<%=rs("id")%>" /></td>
      </tr>
    <tr>
      <td align="right">��ʵ������</td>
      <td height="30"><input name="zsname" type="text" class="input_k" value="<%=rs("zsname")%>" size="30" /></td>
      </tr>
    <tr>
      <td align="right">�Ա�</td>
      <td height="30"><label><input type="radio" name="sex" value="0" <%if rs("sex")=0 then%>checked<%end if%>>����</label>�� 
        <label><input type="radio" name="sex" value="1" <%if rs("sex")=1 then%>checked<%end if%>>Ůʿ</label></td>
      </tr>
      <td align="right">��˾���ƣ�</td>
      <td height="30"><input name="gsname" type="text" class="input_k" value="<%=rs("gsname")%>" size="30" /></td>
      </tr>
    <tr>
      <td align="right">��ϸ��ַ��</td>
      <td height="30"><input name="gsadd" type="text" class="input_k" value="<%=rs("gsadd")%>" size="40" /></td>
      </tr>
    <tr>
      <td align="right">�������룺</td>
      <td height="30"><input name="youbian" type="text" class="input_k" value="<%=rs("youbian")%>" size="20" /></td>
      </tr>
    <tr>
      <td align="right">��ϵ�绰��</td>
      <td height="30"><input name="tel" type="text" class="input_k" value="<%=rs("tel")%>" size="30" /></td>
      </tr>
    <tr>
      <td align="right">�ֻ����룺</td>
      <td height="30"><input name="sj" type="text" class="input_k" value="<%=rs("sj")%>" size="30" /></td>
      </tr>
      <tr>
      <td align="right">��ϵQQ��</td>
      <td height="30"><input name="qq" type="text" class="input_k" value="<%=rs("qq")%>" size="30" /></td>
      </tr>
    <tr>
      <td align="right">�������䣺</td>
      <td height="30"><input name="mail" type="text" class="input_k" value="<%=rs("mail")%>" size="30" /></td>
      </tr>
    <tr>
      <td align="right">��˾��ַ��</td>
      <td height="30"><input name="wz" type="text" class="input_k" value="<%=rs("wz")%>" size="30" /></td>
      </tr>
    
    <tr>
      <td align="right">ע��ʱ�䣺</td>
      <td height="27"><%=rs("data")%></td>
      </tr>
    <tr>
      <td align="right">ע��IP��</td>
      <td height="30"><%=rs("ip")%></td>
      </tr>
    <tr>
      <td align="right">��½������</td>
      <td height="30"><%=rs("dlcs")%></td>
      </tr>
    <tr>
      <td align="right">����½��</td>
      <td height="30"><%=rs("dldata")%></td>
      </tr>
    <tr>
      <td height="27" align="right">&nbsp;</td>
      <td><label>
        <input name="button" type="submit" class="input" id="button" value="��������" />
      </label></td>
      </tr>
  </table></form>
<%rs.close
set rs=nothing%>
</div>
<%end if
if request.querystring("action")="password" then%>
<div  class="ubox">
<div class="pronav"><h1>�޸�����</h1></div>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:20px auto; padding-left:5px">
<tr>
   <td><%if session("username")<>"" then%>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rs.open exec,conn,1,1 %>
<form action="?action=password&id=<%=rs("id")%>&xiugai=pass" method="post" name="add">
<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#D8D8D8" style="border-collapse: collapse;line-height:25px; text-indent:3px">
    <tr>
      <td width="19%" height="30">��½�ʺţ�</td>
      <td width="38%" colspan="2"><%=rs("useradmin")%><input name="id" type="hidden" value="<%=rs("id")%>" /></td>
      <td width="43%">&nbsp;</td>
    </tr>
    <tr>
      <td height="30">��½���룺</td>
      <td colspan="2"><input name="userpassword" type="text" class="input_k" size="30" /></td>
      <td>6-12λ,��ĸ�����ֻ��»���</td>
    </tr>
    <tr>
      <td height="30">�ظ����룺</td>
      <td colspan="2"><input name="userpassword2" type="text" class="input_k" size="30" /></td>
      <td>ͬ��</td>
    </tr>
    <tr>
      <td height="30">&nbsp;</td>
      <td colspan="2"><label><input name="button" type="submit" class="input" id="button" value="�޸�����" /></label></td>
      <td>&nbsp;</td>
    </tr>
  </table></form>
  <%else
response.write ("<div class=""on_login"">��δ��½�����ȵ�½</div>")
response.write ("<div class=""on_loginandRegister"">")
response.write ("<samp>��ע���Ա<a href=""login.asp"">���˵�½</a></samp> <samp>��û���ʺŵ���˴�<a href=""Register.asp"">����ע��</a></samp></div>")
end if%>
</td>
</tr>
</table>
</div>
<% End If %>
</div>
</div>
   
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_U">��Ա���� <span>User Nav</span></div>
<ul class="list_2">
<%=zych_user_nav%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>
<div id="light" class="Order_tc">
<span class="ortitle">����֧����ʾ</span>
<div class="up-box">
<span class="icon"><img src="../images/working.gif"></span>
<div class="r-txt">֧�����ǰ���벻Ҫ�رմ�֧����֤���ڡ�<br clear="none">֧����ɺ��������֧�������������水ť��</div>
</div> 
<div class="lay-btn">
<a href="/User" class="btn92" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">֧����������</a>
<a href="/User/my_orders.asp" class="btn92s" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'" shape="rect">֧�����</a>
</div>
</div> 
<div id="fade" class="Order_bj"> </div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<%if request("xiugai")="info" then'�޸�����
id=request("id")
zsname=request.form("zsname")
sex=request.form("sex")
gsname=request.form("gsname")
gsadd=request.form("gsadd")
youbian=request.form("youbian")
tel=request.form("tel")
qq=Replace(Trim(Request.Form("qq")),"'","''")
sj=request.form("sj")
mail=request.form("mail")
wz=request.form("wz")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
if not isnumeric(qq) then
Response.Write "<script>alert('QQ��ֻ��Ϊ���֣�');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from [user] where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if
rs("zsname")=zsname
rs("sex")=sex
rs("da")=da
rs("gsname")=gsname
rs("gsadd")=gsadd
rs("youbian")=youbian
rs("tel")=tel
rs("qq")=qq
rs("sj")=sj
rs("mail")=mail
rs("wz")=wz
rs.update 
rs.close 
response.write "<script>alert('���ϸ��³ɹ���');window.location.href='?action=admin';</script>" 
end if
 
if request("xiugai")="pass" then'�޸�����
dim u,i,letters,id,userpassword,userpassword2
id=request("id")
userpassword=request.form("userpassword")
userpassword2=request.form("userpassword2")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	if userpassword="" or userpassword2="" then
response.write "<script>alert('���벻��Ϊ��!!');history.go(-1);</script>"  
response.end 
end if
if userpassword<>userpassword2 then 
response.write "<script>alert('�����������벻һ��,����������!');history.go(-1);</script>"  
response.end 
end if
letters="0123456789abcdefghijklmnopqrstuvwxyz" 
userpassword=Lcase(trim(Request.Form("userpassword"))) 
for i=1 to len(userpassword) 
u=mid(userpassword,i,1) 
if Instr(letters,u)=0 then 
response.write "<script>alert('��½����ֻ������ĸ�����ּ��»������!');history.go(-1);</script>" 
response.end 
end if 
next 
if len(userpassword)<6 or len(userpassword)>20 then   
response.write "<script>alert('�������Ϊ6��20λ!');history.go(-1);</script>" 
response.end 
end if 
	SQL="Select * from [user] where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
rs("userpassword")=md5(request.form("userpassword"))
rs.update 
rs.close 
response.write "<script>alert('�����޸ĳɹ���');window.location.href='"&Dir&"User/?action=admin';</script>" 
end if

if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from orders where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰ����ɾ���ɹ���');window.location.href='"&Dir&"User/?action=admin';</script>"
end if 
%> 
