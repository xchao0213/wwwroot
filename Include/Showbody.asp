<!--#include file="conn.asp"-->
<!--#include file="config.asp" -->
<!--#include file="Sql.Asp" -->
<!--#include file="zych_sql.asp"--> 
 <%
'����Ƶ�����ص�ַ
if request("action")="dowurl" then
id=request.QueryString("id") 
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from download where id="&id
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
sql="update download set js=js+1 where id="&id&" and js is not null" '���ش���
conn.execute(sql) 
dowurl=rs("url")
if rs("qx")=9 and session("username")="" then
	Response.Write("<script>alert('��������Ȩ�ޣ����½��Ա');history.go(-1);</script>")
else
	if dowurl="" or dowurl="#" then 
	Response.Write("<script>alert('�������ص�ַ');history.go(-1);</script>")
	else
	response.redirect ""&rs("url")&"" 
	end if
end if
rs.close 
set rs=nothing 
end if
'����JS���URL
if request("action")="adurl" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
response.write "<script>alert('��������');window.location.href='/index.asp';</script>"
Response.End()
end if
exec="select * from ad where id="&id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
sql="update ad set hit=hit+1 where id="& request.QueryString("id") '�������
conn.execute(sql)
adurl=rs("url")
	if adurl="" or adurl="#" then 
	response.write "<script>alert('��ͼ���ӵ�ַΪ�գ�ȷ���󷴻ر�վ��ҳ��');window.location.href='/index.asp';</script>" 
	else
	response.redirect ""&rs("url")&"" 
	end if
rs.close 
set rs=nothing 
end if
%>