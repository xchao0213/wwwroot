<!--#include file="../../Include/inc.asp" -->
<!--#include file="../../Include/Conn.Asp" -->
<!--#include file="../Admin_htmlconfig.Asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../images/style.css" rel="stylesheet" type="text/css" />
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="../images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">正在生成静态</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF" style="height:50px; padding-left:10px">
<span style="width:300px; height:20px;">
<img src="../Images/bg_list.gif" width="5" height="20" id="bar_img" name="bar_img" align="absmiddle"></span>
<span id="bar_txt2" name="bar_txt2" style="font-size:12px; color:#FF0000"></span><span id="bar_txt1" name="bar_txt1" style="font-size:12px">0</span><span style="font-size:12px">%</span>
</td>
</tr>
</table>
<%
Call HtmlCasefl
Call HtmlCase
conn.Close
Set conn = Nothing
%>