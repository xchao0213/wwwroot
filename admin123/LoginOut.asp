<!--#include file="../Include/Conn.asp" -->  
<%
session.Abandon()
	Response.Write("<script language=""JavaScript"">alert(""��ȫ�˳���̨����ϵͳ��"");</script>")
	response.write "<Meta http-equiv='refresh' content='0;URL=login.asp'>"
%>