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
<div class="bar"><b>��Ա���� <i>User</i></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>User/">��Ա����</a> ��Ա��½</span></div>
<div class="ulogin">
<div class="login_ad">
<h1>ϸ�������ǵķ���ר�������ǵ�רҵ</h1>
<h2>Careful in our service, we concentrate on our professional</h2>
</div>
<%if session("username")<>"" then
response.redirect ""&Dir&"User/?action=admin"
end if%>
<div class="login_box">
<form method="post" action="?login=ok&Tourl=<%=tourl%>" name="add">
<div class="title"><div id="h1">��Ա��½</div> <div id="h2">&nbsp;</div></div>
<div class="box">
<div class="um"><input name="useradmin" type="text" class="uame" size="22" placeholder="��Ա�˺�" /></div>
<div class="up"><input name="password" type="password" class="uame" size="22" placeholder="��Ա����"/></div>
<div class="zm"><input name="VerifyCode" type="text" class="im"   size="10" placeholder="��֤��"/><img class="yzm" src="safecode.asp?" onClick="this.src+=Math.random()" alt="ͼƬ�����壿������µõ���֤��" /></div>
</div>
<p><a href="Register.asp" target="_blank">û�б�վ�˺ŵ��ע��</a></p>
<p><button type="submit" class="J_Submit" tabindex="5" id="J_SubmitStatic">�ǡ�¼</button></p>
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
response.Write("<script language=javascript>alert('�������½�ʺ�!');history.go(-1)</script>") 
response.end
end if 
if Request("password")=""  then 
response.Write("<script language=javascript>alert('�������½����!');history.go(-1)</script>") 
response.end
end if 
sql="select * from [user] where useradmin='"&useradmin&"' and userpassword='"&password&"'" 
set rs=conn.execute(sql) 
if rs.eof or rs.bof then 
response.Write("<script language=javascript>alert('�ʺ��������!');history.go(-1)</script>")  
response.End
end if
if rs("sh")=0 then
response.Write("<script language=javascript>alert('�Բ��������ʺ���ʱδͨ����ˣ����Ժ��ٳ��Ե�½!');history.go(-1)</script>")  
response.End()
end if
session("username")=rs("id")
sql="update [user] set dlcs=dlcs+1 where id="&session("username") '��½����+1
conn.execute(sql) 
sql="update [user] set dldata=#"&now&"# where id="&session("username")  '��¼��½ʱ��
conn.execute(sql) 
Response.Write("<script language=""JavaScript"">alert("""&rs("useradmin")&" ��½�ɹ���"");window.location.href='"&tourl&"';</script>")
end if%>