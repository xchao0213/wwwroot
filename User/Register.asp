<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>会员中心_<%=zych_home%></title>
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
<div class="bar"><b>会员中心 <i>User</i></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>User/">会员中心</a>会员注册</span></div>
<div class="w_650 fl">
<div class="position">
<div  class="ubox">
<div class="pronav"><h1>会员注册</h1></div>
<%if userzc<>0 then
response.Write("对不起!会员注册已关闭!")
else%>
<form action="?user=reg" method="post" name="add">  
<table width="99%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse:collapse; line-height:25px; text-indent:3px; margin:0px 3px;"> 
    <tr>
      <td width="17%" align="right">登陆帐号：</td>
      <td height="30" colspan="2"><label>
       <input name="sh" type="hidden" value="<%=usersh%>" /><input name="useradmin" type="text" class="input_k" size="33" />
        <font color="#FF0000">*</font>字母、数字或下划线</label></td>
      </tr>
    <tr>
      <td align="right">登陆密码：</td>
      <td height="30" colspan="2"><input name="userpassword" type="password" class="input_k" size="33" />
        <font color="#FF0000">*</font>6-12位,字母或数字</td>
      </tr>
    <tr>
      <td align="right">重复密码：</td>
      <td height="30" colspan="2"><input name="userpassword2" type="password" class="input_k" size="33" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">真实姓名：</td>
      <td height="30" colspan="2"><input name="zsname" type="text" class="input_k" size="30" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">性别：</td>
      <td height="30" colspan="2">
        <input name="sex" type="radio" value="0" checked="checked" />
        先生
         <input type="radio" name="sex" value="1" />
         女士 <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">公司名称：</td>
      <td height="30" colspan="2"><input name="gsname" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">详细地址：</td>
      <td height="30" colspan="2"><input name="gsadd" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">邮政编码：</td>
      <td height="30" colspan="2"><input name="youbian" type="text" class="input_k" size="20" /></td>
      </tr>
    <tr>
      <td align="right">联系电话：</td>
      <td height="30" colspan="2"><input name="tel" type="text" class="input_k" size="30" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">联系qq：</td>
      <td height="30" colspan="2"><input name="qq" type="text" class="input_k" size="30" />
        <font color="#FF0000">*</font></td>
      </tr>
    <tr>
      <td align="right">手机号码：</td>
      <td height="30" colspan="2"><input name="sj" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">电子邮箱：</td>
      <td height="30" colspan="2"><input name="mail" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">公司网址：</td>
      <td height="30" colspan="2"><input name="wz" type="text" class="input_k" size="30" /></td>
      </tr>
    <tr>
      <td align="right">验证码：</td>
      <td width="16%" height="30"><input name="VerifyCode" type="text" class="input_k"   size="10" /></td>
      <td width="67%"><img src="<%=dir%>Include/safecode.asp?" alt="图片看不清？点击重新得到验证码" width="70" height="28" style="cursor:hand;" onClick="this.src+=Math.random()" /></td>
      </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td colspan="2"><label>
        <input name="button" type="submit" class="input" id="button" value="立即注册会员" />
      </label></td>
      </tr>
  </table>
</form>
<%end if%>
</div>
</div></div>
   
<div class="w_280 fr">
    <div class="box">
      <div class="title tubiao_U">会员导航 <span>User Nav</span></div>
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
response.write "<script>alert('请填写登陆帐号!');history.go(-1);</script>"  
response.end 
end if
letters="0123456789abcdefghijklmnopqrstuvwxyz" 
useradmin=Lcase(trim(Request.Form("useradmin"))) 
for i=1 to len(useradmin) 
u=mid(useradmin,i,1) 
if Instr(letters,u)=0 then 
response.write "<script>alert('登陆帐号只能由字母、数字及下划线组成!');history.go(-1);</script>" 
response.end 
end if 
next 
if len(useradmin)<2 or len(useradmin)>12 then   
response.write "<script>alert('帐号必须为2至12位!');history.go(-1);</script>" 
response.end 
end if 
if userpassword="" or userpassword2="" then
response.write "<script>alert('请填写登陆密码!!');history.go(-1);</script>"  
response.end 
end if
if userpassword<>userpassword2 then 
response.write "<script>alert('两次密码输入不一致,请重新输入!');history.go(-1);</script>"  
response.end 
end if
letters="0123456789abcdefghijklmnopqrstuvwxyz" 
userpassword=Lcase(trim(Request.Form("userpassword"))) 
for i=1 to len(userpassword) 
u=mid(userpassword,i,1) 
if Instr(letters,u)=0 then 
response.write "<script>alert('登陆密码只能由字母、数字及下划线组成!');history.go(-1);</script>" 
response.end 
end if 
next 
if len(userpassword)<6 or len(userpassword)>20 then   
response.write "<script>alert('密码必须为6至20位!');history.go(-1);</script>" 
response.end 
end if 
if zsname="" then
response.write "<script>alert('请填写真实姓名!');history.go(-1);</script>"  
response.end 
end if
if qq="" then
response.write "<script>alert('QQ号不能为空!');history.go(-1);</script>"  
response.end 
end if
if tel="" then
response.write "<script>alert('联系电话不能为空');history.go(-1);</script>"  
response.end 
end if
if  VerifyCode="" then 
response.Write("<script language=javascript>alert('验证码不能为空!');history.go(-1)</script>") 
response.end
end if 
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('验证码错误!');history.go(-1)</script>")
response.End
end if
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where useradmin='"&useradmin&"'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write "<script>alert('对不起，此登陆帐号名已被注册！请更换其它帐号!');history.go(-1);</script>"  
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
Response.Write "<script>alert('恭喜!注册成功，返回登陆！');window.location.href='login.asp';</script>" 
end if%>
