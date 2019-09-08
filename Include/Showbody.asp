<!--#include file="conn.asp"-->
<!--#include file="config.asp" -->
<!--#include file="Sql.Asp" -->
<!--#include file="zych_sql.asp"--> 
 <%
'下载频道下载地址
if request("action")="dowurl" then
id=request.QueryString("id") 
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from download where id="&id
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
sql="update download set js=js+1 where id="&id&" and js is not null" '下载次数
conn.execute(sql) 
dowurl=rs("url")
if rs("qx")=9 and session("username")="" then
	Response.Write("<script>alert('暂无下载权限，请登陆会员');history.go(-1);</script>")
else
	if dowurl="" or dowurl="#" then 
	Response.Write("<script>alert('暂无下载地址');history.go(-1);</script>")
	else
	response.redirect ""&rs("url")&"" 
	end if
end if
rs.close 
set rs=nothing 
end if
'以下JS广告URL
if request("action")="adurl" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
response.write "<script>alert('参数错误！');window.location.href='/index.asp';</script>"
Response.End()
end if
exec="select * from ad where id="&id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
sql="update ad set hit=hit+1 where id="& request.QueryString("id") '点击次数
conn.execute(sql)
adurl=rs("url")
	if adurl="" or adurl="#" then 
	response.write "<script>alert('此图链接地址为空，确定后反回本站首页！');window.location.href='/index.asp';</script>" 
	else
	response.redirect ""&rs("url")&"" 
	end if
rs.close 
set rs=nothing 
end if
%>