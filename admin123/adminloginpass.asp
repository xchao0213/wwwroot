<!--#include file="../Include/Conn.asp" -->
<!--#include file="../zych_sql.asp"--> 
<!--#include file="md5.Asp"-->
<% 
admin=Replace(request.Form("admin"), "'", "''") 
password=md5(Request("password"))
VerifyCode=request.form("VerifyCode") 
if  VerifyCode="" then 
response.Write("<script language=javascript>alert('验证码不能为空!');history.go(-1)</script>") 
response.end
end if 
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('验证码错误!');history.go(-1)</script>")
response.End
end if
sql="select * from admin where admin='"&admin&"' and password='"&password&"'" 
set rs=conn.execute(sql) 
if rs.eof or rs.bof then 
response.Write("<script language=javascript>alert('帐号密码错误!');history.go(-1)</script>")
else 
session("admin")=rs("id") 
session("qx")=int(rs("qx"))
session.timeout=60 '登陆会话时间，60分钟后自动退出。
sql="update admin set dlcs=dlcs+1 where id=" & session("admin") '登陆次数
conn.execute(sql) 
sql="update admin set dldata=#" & date &" " &Time& "# where id=" & session("admin")  '记录登陆时间
conn.execute(sql) 
%>
<%
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
%>
<%response.redirect "index.asp" 
end if 
%>   
