<!--#include file="../Include/Conn.asp" -->  
<%
session.Abandon()
Response.Write("<script language=""JavaScript"">alert(""您已安全退出系统！"");</script>")
response.write "<Meta http-equiv='refresh' content='0;URL=../index.asp'>"
%>