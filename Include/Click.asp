<!--#include file="conn.asp" -->
<!--#include file="config.asp" -->
<!--#include file="Sql.Asp" -->
<!--#include file="zych_sql.asp"--> 
<%
'文章点击数
if request("click")="News" then
set rs = server.createobject("adodb.recordset")
curid=request("ID")
sql = "select * from News where ID= " + cstr(curid) 
rs.open sql,conn,1,1 
sql = "update News SET hit=hit+ 1 where ID ="&curid 
conn.execute sql 
Response.Write("document.writeln("""&rs("hit")&""");")
end if

if request("click")="Culture" then
set rs = server.createobject("adodb.recordset")
curid=request("ID")
sql = "select * from culture where ID= " + cstr(curid) 
rs.open sql,conn,1,1 
sql = "update culture SET hit=hit+ 1 where ID ="&curid 
conn.execute sql 
Response.Write("document.writeln("""&rs("hit")&""");")
end if

if request("click")="Fengfa" then
set rs = server.createobject("adodb.recordset")
curid=request("ID")
sql = "select * from fengfa where ID= " + cstr(curid) 
rs.open sql,conn,1,1 
sql = "update fengfa SET hit=hit+ 1 where ID ="&curid 
conn.execute sql 
Response.Write("document.writeln("""&rs("hit")&""");")
end if


'产品点击数
if request("click")="Products" then
set rs = server.createobject("adodb.recordset")
curid=request("ID")
sql = "select * from Products where ID= " + cstr(curid) 
rs.open sql,conn,1,1 
sql = "update Products SET hit=hit+ 1 where ID ="&curid 
conn.execute sql 
Response.Write("document.writeln("""&rs("hit")&""");")
end if
'视频点击数
if request("click")="video" then
set rs = server.createobject("adodb.recordset")
curid=request("ID")
sql = "select * from video where ID= " + cstr(curid) 
rs.open sql,conn,1,1 
sql = "update video SET hit=hit+ 1 where ID ="&curid 
conn.execute sql 
Response.Write("document.writeln("""&rs("hit")&""");")
end if
'相册点击数
if request("click")="photo" then
set rs = server.createobject("adodb.recordset")
curid=request("ID")
sql = "select * from photo where ID= " + cstr(curid) 
rs.open sql,conn,1,1 
sql = "update photo SET hit=hit+ 1 where ID ="&curid 
conn.execute sql 
Response.Write("document.writeln("""&rs("hit")&""");")
end if


if request("click")="Visits" then
set rs = server.createobject("adodb.recordset")
sql="update config set js=js+1 where js is not null" 
conn.execute(sql) 
sql = "select * from config"
rs.open sql,conn,1,1 
Response.Write("document.writeln(""您是第:"&rs("js")&"位访问者"");")
end if%>
