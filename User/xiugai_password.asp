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
<div class="bar"><b>会员中心 <i>User</i></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>User/">修改密码</a></span></div>
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
      <td width="19%" height="30">登陆帐号：</td>
      <td width="38%" colspan="2"><%=rs("useradmin")%><input name="id" type="hidden" value="<%=rs("id")%>" /></td>
      <td width="43%">&nbsp;</td>
    </tr>
    <tr>
      <td height="30">登陆密码：</td>
      <td colspan="2"><input name="userpassword" type="text" class="input_k" size="30" /></td>
      <td>6-12位,字母、数字或下划线</td>
    </tr>
    <tr>
      <td height="30">重复密码：</td>
      <td colspan="2"><input name="userpassword2" type="text" class="input_k" size="30" /></td>
      <td>同上</td>
    </tr>
    <tr>
      <td height="30">&nbsp;</td>
      <td colspan="2"><label><input name="button" type="submit" class="input" id="button" value="修改密码" /></label></td>
      <td>&nbsp;</td>
    </tr>
  </table></form>
  <%else
response.write ("<div class=""on_login"">您未登陆！请先登陆</div>")
response.write ("<div class=""on_loginandRegister"">")
response.write ("<samp>已注册会员<a href=""login.asp"">请点此登陆</a></samp> <samp>还没有帐号点击此处<a href=""Register.asp"">立即注册</a></samp></div>")
end if%>
</td>
</tr>
</table>
</div>
</div>
   
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
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
if not isnumeric(qq) then
Response.Write "<script>alert('QQ号只能为数字！');history.go(-1);</script>" 
Response.End()
end if
SQL="Select * from [user] where id="&id
set rs=server.createobject("adodb.recordset")
rs.open SQL,conn,1,3
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
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
response.write "<script>alert('资料更新成功！');window.location.href='index.asp';</script>" 
end if%> 
