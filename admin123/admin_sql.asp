<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"--> 
<%call chkAdmin(6)
Server.ScriptTimeout	=500						
URL= Request.ServerVariables("URL")
Action= Request("Action")
	Select Case Action
		Case "Del"
			Call Delip()
		Case "lock"
			Call lockIP()
		Case "unlock"
			Call UnLockip()
		Case "Logout"
			Call Logout()
		Case "config"
			Call config()
		Case "saveconfig"
			Call saveconfig()
		Case Else 
			Call Main()
	end Select


Sub Login()
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Sql��ע�����</title>
<link href="images/style.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
End Sub
Sub Main()
	Call header()
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
<body> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">SQLע�����</div> </td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="4" cellspacing="0">
<tr>
<td height="30" align="center" bgcolor="#F1F5F8">���</td>
<td align="center" bgcolor="#F1F5F8">�����ɣ�</td>
<td align="center" bgcolor="#F1F5F8">��ǰ״̬</td>
<td align="center" bgcolor="#F1F5F8">�Ƿ�����</td>
<td align="center" bgcolor="#F1F5F8">����ҳ��</td>
<td align="center" bgcolor="#F1F5F8">����ʱ��</td>
<td align="center" bgcolor="#F1F5F8">�ύ��ʽ</td>
<td align="center" bgcolor="#F1F5F8">�ύ����</td>
<td align="center" bgcolor="#F1F5F8">�ύ����</td>
</tr>
      <tr align="center" bgcolor="#FFFFFF">
<%
sql="select * from SqlIn order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write ("<td align=""center"" colspan=""9"">����ע���¼...</td>")
else
'��ҳ��ʵ�� 
listnum=20
Rs.pagesize=listnum
page=Request("page")
if (page-Rs.pagecount) > 0 then
page=rs.pagecount
elseif page = "" or page < 1 then
page = 1
end if
rs.absolutepage=page
'��ŵ�ʵ��
j=rs.recordcount
j=j-(page-1)*listnum
i=0
nn=request("page")
if nn="" then
n=0
else
nn=nn-1
n=listnum*nn
end if%>

<%do while not rs.eof and i<listnum
n=n+1%>
      <form action="?del=ok" method="post" name="check" id="check">
        <tr align="center" bgcolor="#FFFFFF" height="22">
          <td class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" />
              <%=n%></td>
          <td class="td"><%=rs("SqlIn_IP")%></td>
          <td class="td">
		  <%
		  if rs("Kill_ip")=true then 
			response.write "<font color='red'>������</font>"
		 else
			response.write "<font color='green'>�ѽ���</font>"
		 end if%>
		 </td>
          <td class="td">
		  <%if rs("Kill_ip")=true then 
			response.write "<a href="&URL&"?action=unlock&id="&rs("id")&" style=""color:#FF0000"">����IP</a>"
		 else
			response.write "<a href="&URL&"?action=lock&id="&rs("id")&" style=""color:#006600"">����IP</a>"
		end if
	      %>
          </td>
          <td class="td"><%=rs("SqlIn_WEB")%></td>
          <td class="td"><%=rs("SqlIn_TIME")%></td>
          <td class="td"><%=rs("SqlIn_FS")%></td>
          <td class="td"><%=rs("SqlIn_CS")%></td>
          <td class="td"><%=N_Replace(rs("SqlIn_SJ"))%></td>
        </tr>
<%rs.movenext 
i=i+1 
j=j-1
loop%>
        <tr bgcolor="#FFFFFF">
          <%filename=URL%>
          <td colspan="9" align="right">��<%=Rs.recordcount%> ����¼&nbsp;&nbsp;<%=listnum%> ����¼/ҳ&nbsp;&nbsp;�� <%=rs.pagecount%> ҳ
            <% if page=1 then %>
             <%else%>
             <a href="<%=filename%>"><strong>|&lt;&lt;</strong></a> <a href="<%=filename%>?page=<%=page-1%>"><strong>&lt;&lt;</strong></a> <a href="<%=filename%>?page=<%=page-1%>"><b>[<%=page-1%>]</b></a>
             <%end if%>
            <% if rs.pagecount=1 then%>
            <%else%>
            <b>[<%=page%>]</b>
            <%end if%>
              <% if rs.pagecount-page <> 0 then %>
             <a href="<%=filename%>?page=<%=page+1%>"><b>[<%=page+1%>]</b></a> <a href="<%=filename%>?page=<%=page+1%>"><strong>&gt;&gt;</strong></a> <a href="<%=filename%>?page=<%=rs.pagecount%>"><strong>&gt;&gt;|</strong></a>
              <%end if%>
              <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0" />
            ȫѡ
            <input type="submit" name="action" value="ɾ��" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" class="btn"/>
          </td>
          <%end if%>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
<%
Call footer()
end Sub
sub config()
	Call header()
	Set rsinfo=conn.execute("select * from sqlconfig")
	N_In		= rsinfo("N_In")
	Kill_IP		= rsinfo("Kill_IP")			
	WriteSql	= rsinfo("WriteSql")		
	alert_url	= rsinfo("alert_url")
	alert_info	= rsinfo("alert_info")
	kill_info	= rsinfo("kill_info")
	N_type		= rsinfo("N_type")
	Sec_Forms	= rsinfo("Sec_Forms")
	Sec_Form_open = rsinfo("Sec_Form_open")
	rsinfo.close
	Set rsinfo=Nothing 
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">SQL��ע������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      
      <tr>
        <td><table width="100%"  border="0" cellpadding="5" cellspacing="0" class="stable" style="border-collapse:collapse;">
            <form name="form" method="post" action="<%=url%>?action=saveconfig">
         <tr >
                <td width="18%"  align="right"  class="td">��Ҫ���˵Ĺؼ��֣�</td>
                <td width="82%" class="td">&nbsp;
                    <input name="N_In" type="text" value="<%=N_In%>" id="r_str" style=" " size="50">
                  ��&quot;|&quot;�ֿ�</td>
              </tr>
              <tr>
                <td align="right" class="td">�Ƿ��¼��������Ϣ��</td>
                <td class="td">&nbsp;
                    <select name="WriteSql" id="r_kill">
                      <option value="1" <%if WriteSql=1 Then response.write "selected"%>>��</option>
                      <option value="0" <%if WriteSql=0 Then response.write "selected"%>>��</option>
                  </select></td>
              </tr>
                <tr >
                <td align="right" class="td">�Ƿ���������IP��</td>
                <td class="td">&nbsp;
                    <select name="Kill_IP" id="r_kill">
                      <option value="1" <%if Kill_IP=1 Then response.write "selected"%>>��</option>
                      <option value="0" <%if Kill_IP=0 Then response.write "selected"%>>��</option>
                  </select></td>
              </tr>
              <tr>
                <td align="right" class="td">�Ƿ����ð�ȫҳ�棺</td>
                <td class="td">&nbsp;
                    <select name="Sec_Form_open" id="r_kill">
                      <option value="1" <%if Sec_Form_open=1 Then response.write "selected"%>>��</option>
                      <option value="0" <%if Sec_Form_open=0 Then response.write "selected"%>>��</option>
                    </select>
                  ����������ܣ��������ȷ�ϴ�ҳ��������ˣ���ȷ���԰�ȫûӰ�죡 </td>
              </tr>
                <tr >
                <td align="right" class="td">����Ϊ��ȫ��ҳ�棺</td>
                <td class="td">&nbsp;
                    <input name="Sec_Forms" type="text" value="<%=Sec_Forms%>" id="r_str" style=" " size="50">
                  ��&quot;|&quot;�ֿ�</td>
              </tr>
              <tr>
                <td align="right" class="td">�����Ĵ���ʽ��</td>
                <td class="td">&nbsp;
                    <select name="N_type" id="r_kill">
                      <option value="1" <%if N_type=1 Then response.write "selected"%>>ֱ�ӹر���ҳ</option>
                      <option value="2" <%if N_type=2 Then response.write "selected"%>>�����ر�</option>
                      <option value="3" <%if N_type=3 Then response.write "selected"%>>��ת��ָ��ҳ��</option>
                      <option value="4" <%if N_type=4 Then response.write "selected"%>>�������ת</option>
                  </select></td>
              </tr>
                 <tr >
                <td align="right" class="td">�������תUrl��</td>
                <td class="td">&nbsp;
                   <input name="alert_url" type="text" value="<%=alert_url%>" id="r_str" style=" " size="30"></td>
              </tr>
              <tr>
                <td align="right" class="td">������ʾ��Ϣ��</td>
                <td class="td">&nbsp;
                  <textarea name="alert_info" cols="45" rows="4" id="alert_info" style=";  "><%=alert_info%></textarea>                  
                  \n\n����</td>
              </tr>
                 <tr >
                <td align="right" class="td">��ֹ������ʾ��Ϣ��</td>
                <td class="td">&nbsp;
                    <textarea name="kill_info" cols="45" rows="4" id="r_err2" style=";  "><%=kill_info%></textarea>
                  \n\n���� </td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td><input name="enter_3" type="submit" id="enter_3" value="��������"  class="btn"></td>
              </tr>
            </form>
        </table>    	</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
<%
	Call footer()
end Sub
Sub header()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Sql��ע��ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" type="text/css" rel="stylesheet" />
<style type="text/css">
<!--
body,td,th {font-size: 12px;}
a {font-size: 12px;}
body {margin-left: 0px;margin-top: 0px;margin-right: 0px;margin-bottom: 0px;}
a:link {color: #333333;text-decoration: none;}
a:visited {text-decoration: none;color: #333333;}
a:hover {text-decoration: underline;color: #FF3300;}
a:active {text-decoration: none;color: #333333;}
-->
</style>
<%
End Sub 
	sub footer()
%>
<%
end Sub
Sub Delip()
dim id 
conn.execute("delete from SqlIn where id in ( " & id & ")")
Response.Redirect URL
End sub

Sub Lockip()
id = clng(request("id"))
conn.execute("update SqlIn set Kill_ip=true where id="&id)
Response.Redirect URL
End sub

Sub UnLockip()
id = clng(request("id"))
conn.execute("update SqlIn set Kill_ip=False where id="&id)
Response.Redirect URL
End sub

Sub Logout()
	Session("AdminPassWord")="NUll"
	Response.Redirect URL
End Sub

Sub saveconfig
	N_In		=replace(request.form("N_In"),"'","''")
	Kill_IP		=request.form("Kill_IP")			
	WriteSql	=request.form("WriteSql")		
	alert_url	=request.form("alert_url")
	alert_info	=request.form("alert_info")
	kill_info	=request.form("kill_info")
	N_type		=request.form("N_type")
	Sec_Forms	=request.form("Sec_Forms")
	Sec_Form_open=request.form("Sec_Form_open")

	sql="update sqlconfig set N_In='"&N_In&"',Kill_IP="&Kill_IP&",WriteSql="&WriteSql&",alert_url='"&alert_url&"',alert_info='"&alert_info&"',kill_info='"&kill_info&"',N_type="&N_type&",Sec_Forms='"&Sec_Forms&"',Sec_Form_open="&Sec_Form_open&""
		Response.Write "<script language='javascript'>alert('����ɹ���');document.location.href('admin_sql.asp?Action=config');</script>"
	conn.execute(sql)
	Application.Lock
	set Application("Neeao_config_info")=nothing
	Application.unlock
	Call main()
End Sub 

Function N_Replace(N_urlString)
	N_urlString = Replace(N_urlString,"'","''")
    N_urlString = Replace(N_urlString, ">", "&gt;")
    N_urlString = Replace(N_urlString, "<", "&lt;")
    N_Replace = N_urlString
End Function
%>
<%
if Request.QueryString("del")="ok" then
if Request("id")="" then
Response.Write "<script>alert('��ѡ��Ҫɾ���ļ�¼!');window.location.href='admin_sql.asp';</script>" 
response.end()
end if
dim sql
sql="delete from [SqlIn] where id in ("&Request("id")&")"
conn.Execute ( sql )
conn.close
set conn=nothing
Response.Write "<script>alert('����ɾ���ɹ�!');window.location.href='admin_sql.asp';</script>" 
end if
%>
