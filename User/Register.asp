<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
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
<div class="bar"><b>��Ա���� <i>User</i></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>User/">��Ա����</a>��Աע��</span></div>
<div class="w_650 fl">
<div class="position">
<div  class="ubox">
<div class="pronav"><h1>��Աע��</h1></div>
<%if userzc<>0 then
response.Write("�Բ���!��Աע���ѹر�!")
else%>
<form action="?user=reg" method="post" name="add">  
<table width="99%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse:collapse; line-height:25px; text-indent:3px; margin:0px 3px;"> 
    <tr>
      <td width="17%" align="right">��½�ʺţ�</td>
      <td height="30" colspan="2"><label>
       <input name="sh" type="hidden" value="<%=usersh%>" /><input name="useradmin" type="text" class="input_k" size="33" />
        <font color="#FF0000">*</font>��ĸ�����ֻ��»���</label></td>
      </tr>
    <tr>
      <td align="right">��½���룺</td>
      <td height="30" colspan="2"><input name="userpassword" type="password" class="input_k" size="33" />
        <font color="#FF0000">*</font>6-12λ,��ĸ������</td>
      </tr>
    <tr>
      <td align="right">�ظ����룺</td>
      <td height="30" colspan="2"><input name="userpassword2" type="password" class="input_k" size="33" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">��ʵ������</td>
      <td height="30" colspan="2"><input name="zsname" type="text" class="input_k" size="30" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">�Ա�</td>
      <td height="30" colspan="2">
        <input name="sex" type="radio" value="0" checked="checked" />
        ����
         <input type="radio" name="sex" value="1" />
         Ůʿ <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">��˾���ƣ�</td>
      <td height="30" colspan="2"><input name="gsname" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">��ϸ��ַ��</td>
      <td height="30" colspan="2"><input name="gsadd" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">�������룺</td>
      <td height="30" colspan="2"><input name="youbian" type="text" class="input_k" size="20" /></td>
      </tr>
    <tr>
      <td align="right">��ϵ�绰��</td>
      <td height="30" colspan="2"><input name="tel" type="text" class="input_k" size="30" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">��ϵqq��</td>
      <td height="30" colspan="2"><input name="qq" type="text" class="input_k" size="30" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">�ֻ����룺</td>
      <td height="30" colspan="2"><input name="sj" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">�������䣺</td>
      <td height="30" colspan="2"><input name="mail" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">��˾��ַ��</td>
      <td height="30" colspan="2"><input name="wz" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">��֤�룺</td>
      <td width="16%" height="30"><input name="VerifyCode" type="text" class="input_k"   size="10" /></td>
      <td width="67%"><img src="<%=dir%>Include/safecode.asp?" alt="ͼƬ�����壿������µõ���֤��" width="70" height="28" style="cursor:hand;" onClick="this.src+=Math.random()" /></td>
      </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td colspan="2"><label>
        <input name="button" type="submit" class="input" id="button" value="����ע���Ա" />
      </label></td>
      </tr>
  </table>
</form>
<%end if%>
</div>
</div></div>
   
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
<%
if request("user")="reg" then
ip=request.servervariables("remote_addr")
useradmin=replace(trim(request("useradmin")),"'","") 
userpassword=replace(trim(Request("userpassword")),"'","") 
userpassword2=replace(trim(Request("userpassword2")),"'","") 
zsname=request.form("zsname")
sex=request.form("sex")
wen=request.form("wen")
da=request.form("da")
gsname=request.form("gsname")
youbian=request.form("youbian")
gsadd=request.form("gsadd")
tel=request.form("tel")
qq=request.form("qq")
sj=request.form("sj")
mail=request.form("mail")
wz=request.form("wz")
sh=request.form("sh")
VerifyCode=request.form("VerifyCode")
if useradmin="" then
response.write "<script>alert('����д��½�ʺ�!');history.go(-1);</script>"  
response.end 
end if
letters="0123456789abcdefghijklmnopqrstuvwxyz" 
useradmin=Lcase(trim(Request.Form("useradmin"))) 
for i=1 to len(useradmin) 
u=mid(useradmin,i,1) 
if Instr(letters,u)=0 then 
response.write "<script>alert('��½�ʺ�ֻ������ĸ�����ּ��»������!');history.go(-1);</script>" 
response.end 
end if 
next 
if len(useradmin)<2 or len(useradmin)>12 then   
response.write "<script>alert('�ʺű���Ϊ2��12λ!');history.go(-1);</script>" 
response.end 
end if 
if userpassword="" or userpassword2="" then
response.write "<script>alert('����д��½����!!');history.go(-1);</script>"  
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
if zsname="" then
response.write "<script>alert('����д��ʵ����!');history.go(-1);</script>"  
response.end 
end if
if qq="" then
response.write "<script>alert('QQ�Ų���Ϊ��!');history.go(-1);</script>"  
response.end 
end if
if tel="" then
response.write "<script>alert('��ϵ�绰����Ϊ��');history.go(-1);</script>"  
response.end 
end if
if  VerifyCode="" then 
response.Write("<script language=javascript>alert('��֤�벻��Ϊ��!');history.go(-1)</script>") 
response.end
end if 
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('��֤�����!');history.go(-1)</script>")
response.End
end if
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where useradmin='"&useradmin&"'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write "<script>alert('�Բ��𣬴˵�½�ʺ����ѱ�ע�ᣡ����������ʺ�!');history.go(-1);</script>"  
response.end 
end if
rs.addnew
rs("useradmin")=useradmin
rs("userpassword")=md5(request.form("userpassword"))
rs("zsname")=zsname
rs("sex")=sex
rs("ip")=ip
rs("gsname")=gsname
rs("youbian")=youbian
rs("gsadd")=gsadd
rs("tel")=tel
rs("qq")=qq
rs("sj")=sj
rs("mail")=mail
rs("wz")=wz
rs("sh")=sh
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('��ϲ!ע��ɹ������ص�½��');window.location.href='login.asp';</script>" 
end if%>
