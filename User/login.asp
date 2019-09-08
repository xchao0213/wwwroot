<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<%tourl=session("Tourl")%>
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
<div class="bar"><b>会员中心 <i>User</i></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>User/">会员中心</a> 会员登陆</span></div>
<div class="ulogin">
<div class="login_ad">
<h1>细心于我们的服务，专心于我们的专业</h1>
<h2>Careful in our service, we concentrate on our professional</h2>
</div>
<%if session("username")<>"" then
response.redirect ""&Dir&"User/?action=admin"
end if%>
<div class="login_box">
<form method="post" action="?login=ok&Tourl=<%=tourl%>" name="add">
<div class="title"><div id="h1">会员登陆</div> <div id="h2">&nbsp;</div></div>
<div class="box">
<div class="um"><input name="useradmin" type="text" class="uame" size="22" placeholder="会员账号" /></div>
<div class="up"><input name="password" type="password" class="uame" size="22" placeholder="会员密码"/></div>
<div class="zm"><input name="VerifyCode" type="text" class="im"   size="10" placeholder="验证码"/><img class="yzm" src="safecode.asp?" onClick="this.src+=Math.random()" alt="图片看不清？点击重新得到验证码" /></div>
</div>
<p><a href="Register.asp" target="_blank">没有本站账号点击注册</a></p>
<p><button type="submit" class="J_Submit" tabindex="5" id="J_SubmitStatic">登　录</button></p>
</form>
</div>
</div>
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<% if Request.QueryString("login")="ok" then
useradmin=Replace(request.Form("useradmin"), "'", "''") 
password=md5(Request("password"))
if useradmin=""  then 
response.Write("<script language=javascript>alert('请输入登陆帐号!');history.go(-1)</script>") 
response.end
end if 
if Request("password")=""  then 
response.Write("<script language=javascript>alert('请输入登陆密码!');history.go(-1)</script>") 
response.end
end if 
sql="select * from [user] where useradmin='"&useradmin&"' and userpassword='"&password&"'" 
set rs=conn.execute(sql) 
if rs.eof or rs.bof then 
response.Write("<script language=javascript>alert('帐号密码错误!');history.go(-1)</script>")  
response.End
end if
if rs("sh")=0 then
response.Write("<script language=javascript>alert('对不起，您的帐号暂时未通过审核！请稍候再尝试登陆!');history.go(-1)</script>")  
response.End()
end if
session("username")=rs("id")
sql="update [user] set dlcs=dlcs+1 where id="&session("username") '登陆次数+1
conn.execute(sql) 
sql="update [user] set dldata=#"&now&"# where id="&session("username")  '记录登陆时间
conn.execute(sql) 
Response.Write("<script language=""JavaScript"">alert("""&rs("useradmin")&" 登陆成功！"");window.location.href='"&tourl&"';</script>")
end if%>