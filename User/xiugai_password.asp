<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!--#include file="../Include/seeion.asp" -->
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
<div class="bar"><b>��Ա���� <i>User</i></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>User/">�޸�����</a></span></div>
<div class="w_650 fl">
<div class="position ">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:20px auto; padding-left:5px">
<tr>
   <td><%if session("username")<>"" then%>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rs.open exec,conn,1,1 %>
<form action="?id=<%=rs("id")%>&xiugai=ok" method="post" name="add">
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

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<%if request("xiugai")="ok" then
id=request("id")
zsname=request.form("zsname")
sex=request.form("sex")
da=request.form("da")
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
response.write "<script>alert('���ϸ��³ɹ���');window.location.href='index.asp';</script>" 
end if%> 
