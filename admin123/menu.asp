<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>main</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style type="text/css">
<!--
body,td,th {font-size: 12px;font-family: 宋体; padding:0px 5px 0px 10px}
a {font-size: 12px;}
body {margin-left: 0px;margin-top: 3px;margin-right: 0px;margin-bottom: 0px;}
a:link {color:#333333;text-decoration: none;}
a:visited {text-decoration: none;color: #333333;}
a:hover {text-decoration: underline;color: #FF3300;}
a:active {text-decoration: none;color: #333333;}
#admin{width:230px; height:30px; float:left; margin-left:3px;}
#version{width:400px; height:30px; float:right; text-align:right; margin-right:3px}
#Dynamic_box {float:left;text-align:left;overflow:hidden;padding:0px;width:420px; height:30px;font-size:12px;font-family: "宋体";}
#Dynamic_box ul {list-style:none; padding:0px; margin:0px;}
#Dynamic_box li {line-height:30px;color:#FF0000; list-style-type:none;}
#Dynamic_box li a{color:#FF0000; text-decoration:none;}
-->
</style>
<SCRIPT language=javascript src="<%=dir%>js/MSClass.js" type=text/javascript></SCRIPT>
</head>
<body>
<div style="width:100%; height:30px; line-height:30px; background:#F1F5F8; border:1px #999 solid">
<div id="admin"><strong>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from admin where id="&session("admin")&""
rs.open exec,conn,1,1 %> 
<font color="#FF0000"><%=rs("admin")%></font>身份：<%if rs("key")<>0 then
response.Write("<font color=#336699>普通管理员</font>")
else
response.write("<font color=#336699>超级管理员</font>")
end if%>
&nbsp;</strong>&nbsp;官方动态：</div>
<div id="Dynamic_box"><ul id=TextDiv2>
<script src=http://www.shgwzy.com/adzychfile/zych_Dynamic.js></script></ul>
<SCRIPT defer>new Marquee("TextDiv2",0,0.2,500,30,20,4000,500,30) </SCRIPT>
</div>
<div id="version">程序当前版本：<font color="#FF0000">高屋置业企业网站系统<%= zychversion %>商业版 <%=data%></font></div>
</td>
</div>
</body>
</html>