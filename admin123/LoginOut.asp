<!--#include file="../Include/Conn.asp" -->  
<%
session.Abandon()
	Response.Write("<script language=""JavaScript"">alert(""安全退出后台管理系统！"");</script>")
	response.write "<Meta http-equiv='refresh' content='0;URL=login.asp'>"
%>