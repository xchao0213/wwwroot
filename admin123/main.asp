<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%az=""&request.ServerVariables("SERVER_NAME")&""
wz=request.ServerVariables("HTTP_HOST")&request.ServerVariables("URL")
wurl=right(left(wz,instrrev(wz,"/")),(len(left(wz,instrrev(wz,"/")))-len(request.ServerVariables("HTTP_HOST")))-1)%>  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>��������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
</head>
<style type="text/css">
<!--

-->
</style>
<body>

<%If az="127.0.0.1" or az="localhost" or Instr(az,"192.168")>0 Then 
Response.Write("")
else
If wurl="admin/" Then %>
<div class="pop" id="pop" style="display:none;">
<div class="box"><a href="javascript:void(0)" onclick ="document.getElementById('pop').style.display='none'" style="position:absolute;right:0px;color:#000; margin:2px 10px;" title="�رմ���" id="closeBtn">���ر�</a>
<p class="t1">������վ��̨Ĭ��Ŀ¼Ϊ��<a>admin</a>��,Ϊ�˰�ȫ������޸�Ϊ��������</p>
<form id="form" name="form" method="post" action="?admin=ok">
<p style="text-align:center"><input name="admin" type="text" class="value" id="admin" value="" placeholder="���ĺ�̨����"/><input class="button" name="button" type="submit" value="ȷ���޸�" /></p>
</form>
<p class="t2">�޸���ɺ������´򿪺�̨��</p>
</div>
</div>
<script>
(function(){
    setTimeout(function(){
        var obj = document.getElementById("pop");
        obj.style.display ="block";
    },3000);
})();
</script>
<% End If %>
<% if Request.QueryString("admin")="ok" then
newadmin=request.form("admin")
if newadmin=""  then 
response.Write("<script language=javascript>alert('Ŀ¼���Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
set fso=Server.CreateObject("Scripting.FileSystemObject") 
set file=fso.GetFolder(Server.MapPath("../admin")) 
file.name=request.form("admin")
set fso=nothing
Response.Write("<div id=""pop"" style=""display:block;""><div class=""box"">")
Response.Write("<h1>�޸ĳɹ���</h1>")
Response.Write("<p style=""font-size:16px;color:#F00;font-family:'΢���ź�'"">���ĺ�̨��½Ŀ¼�Ѿ��޸�Ϊ��"&newadmin&"��,���μ�</p>")
Response.Write("<p><a style=""font-size:20px;text-align:center;"" href=""/"&newadmin&""" target=""_blank"">[�����Ե��������µĺ�̨]</a></p>")
Response.Write("</div></div>")
end if
End If
if Instr(request.ServerVariables("SERVER_SOFTWARE"),"NetBox")>0 then%>
<div class="pop" id="pip" style="display:none;">
<div class="box"><a href="javascript:void(0)" onclick ="document.getElementById('pip').style.display='none'" style="position:absolute;right:0px;color:#000; margin:2px 10px;" title="�رմ���" id="closeBtn">���ر�</a>
<p class="t2">�ڲ��Ա���վʱ������Ҫȷ����վ�Ƿ���<br />
<a style="color:#F00">����Windows-IIS������ �� �ϴ���FTP�ռ���в���</a></p>
<p>��������������С���ߣ�����С���߻�Ӱ����վ����ʹ�ã����̨�����ϴ�ͼƬ��</p>
<p><img src="../images/NetBox.jpg" width="79" height="66" /></p>
</div>
</div>
<script>
(function(){
    setTimeout(function(){
        var obj = document.getElementById("pip");
        obj.style.display ="block";
    },3000);
})();
</script>
<%End If %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">վ����Ϣ</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
<tr bgcolor="#F1F5F8" >
    <td height="25" colspan="6" class="td">
  <span class="mian_bt">��л��ʹ�ø�����ҵ��ҵ��վ����ϵͳ<%=zychversion%>�棬���֤�ţ�<%=authorized%> DATA:<%=data%></span>
  <div class="mian_body"><ul style="padding-left:25px; margin:0px">
  <li>ʹ�ñ�ϵͳע�����</li>
  <li>�������ɸ�����ҵ��������Ȩ�������ҵ���У�</li>
  <li>��ϵͳΪ��ҵ����δ����Ȩ�Ͻ�ʹ�ã�������Ҫ���ڱ���˾��ϵ�������������Ͷ��ɹ����κε�λ�����˲����޸İ�Ȩ�ͽ��ж������ۣ�</li>
  <li>Ϊ�˺����ܸ��õĽ������վʹ�ù��������������⣬�벻Ҫɾ���磺<%=data%>����İ汾��,�Ա��պ����ʶ��!</li>
  <li>����ʹ�ù��̣��緢��BUG���뼰ʱ��ϵ������ҵ����Ա��QQ��495573637 ������ҵ�ˣ�������˾���ر��л�������ǵ�֧�֣�www.shgwzy.com</li>
  </ul>
  </div>	
  
    </td>
</tr>
<tr >
      <td width="10%" height="25" align="right" class="td">����Ա������</td>
<td width="20%"  class="td"><b><%=conn.execute("select count(*) from admin")(0)%> </b>��</td>
      <td width="8%" align="right"  class="td">��Ʒ������</td>
      <td width="18%"  class="td"><b><%=conn.execute("select count(*) from products")(0)%></b> ��</td>
      <td width="12%" align="right"  class="td">�¶���/ȫ��������</td>
      <td width="32%"  class="td"><b><font color="#FF0000"><%=conn.execute("select count(*) from orders where state=1")(0)%></font>/<%=conn.execute("select count(*) from orders ")(0)%></b> [<a href="admin_orders.asp?action=admin">�鿴</a>]</td>
</tr>
    <tr>
      <td width="10%" height="25" align="right" class="td">����������</td>
      <td class="td"><b><%=conn.execute("select count(*) from news")(0)%></b> ��</td>
      
      <td align="right" class="td">�յ��ļ�����</td>
      <td class="td"><b><%=conn.execute("select count(*) from Resume")(0)%></b> ��</td>
      <td align="right" class="td">������/ȫ�����ԣ�</td>
      <td class="td"><b><font color="#FF0000"><%=conn.execute("select count(*) from ly where hf is null")(0)%></font>/<%=conn.execute("select count(*) from ly ")(0)%></b> [<a href="admin_message.asp?action=admin">�鿴</a>]</td>
    </tr>
    <tr>
      <td width="10%" height="25" align="right" class="td">���ظ�����</td>
      <td class="td"><b><%=conn.execute("select count(*) from download")(0)%></b> ��</td>
      <td align="right" class="td">��Ա������</td>
      <td align="left" class="td"><b><%=conn.execute("select count(*) from [user]")(0)%></b> ��</td>
      <td align="right" class="td">��Ƭ������</td>
      <td align="left" class="td"><b><%=conn.execute("select count(*) from photo")(0)%></b> ��</td>
    </tr>
  <tr>
    <td height="25" align="right" class="td">&nbsp;</td>
    <td class="td">&nbsp;</td>
    <td align="right" class="td">�Ŷ�������</td>
    <td class="td"><b><%=conn.execute("select count(*) from tuandui")(0)%></b> ��</td>
    <td align="right" class="td">��Ƶ������</td>
    <td class="td"><b><%=conn.execute("select count(*) from [video]")(0)%></b> ��</td>
  </tr>
</table></td>
</tr>
</table></div>
       </td>
	  </tr>
</table>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
            <tr>
              <td height="28" background="images/bg_list.gif" bgcolor="#F7F7F7"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��������Ϣ</div></td>
            </tr>
            
            <tr>
              <td bgcolor="#FFFFFF" ><table width="100%" border="0" align="center" cellpadding="4" cellspacing="0" class="stable" >
                
<tr >
                  <td height="25" width="13%" class="td">�������·��:</td>
                <td width="43%"  class="td"><%=Server.MapPath("nowaasp.asp")%></td>
                <td width="10%"  class="td">�������ű����棺</td>
                <td width="34%"  class="td"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
            <tr>
                  <td height="25" width="13%" class="td">վ������·����</td>
                <td class="td"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
                <td class="td">·����Ϣ��</td>
                <td class="td"><%=request.ServerVariables("PATH_INFO")%></td>
            </tr>
              
                <tr>
                  <td height="25" width="13%" class="td">�������IP��ַ��</td>
                  <td class="td"><%=request.ServerVariables("REMOTE_ADDR")%></td>
                  <td class="td">�������ڲ�IP��</td>
                  <td class="td"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
                </tr>
                <tr>
                  <td height="25" width="13%"  class="td">SCRIPT����·����</td>
                  <td  class="td"><%=request.ServerVariables("SCRIPT_NAME")%></td>
                  <td  class="td">������IP��ַ��</td>
                  <td  class="td"><%=request.ServerVariables("SERVER_NAME")%></td>
                </tr>
              <tr>
                  <td height="25" width="13%" class="td">�������˿ڣ�</td>
                <td class="td"><%=request.ServerVariables("SERVER_PORT")%></td>
                <td class="td">Э�����ƺͰ汾��</td>
                <td class="td"><%=request.ServerVariables("SERVER_PROTOCOL")%></td>
              </tr>
                 <tr>
                  <td height="25" width="13%" class="td">������IIS�汾��</td>
                  <td class="td"><%=request.ServerVariables("SERVER_SOFTWARE")%></td>
                  <td class="td">�ű���ʱʱ�䣺</td>
                  <td class="td"><%=Server.ScriptTimeout%>��</td>
                </tr>
               <tr>
                  <td height="25" width="13%" class="td">����������ϵ�y��</td>
                 <td  class="td"><%=Request.ServerVariables("OS")%>&nbsp;</td>
                 <td  class="td">������CPU������</td>
                 <td  class="td"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
               </tr>
               <tr>
                 <td height="25" class="td">FSO�ı���д��</td>
                 <td  class="td">
				 <% if IsObjInstalled("Scripting.FileSystemObject") = False Then
				 response.Write("<font color=#aaaaaa><b>��</b></font>")
                 Else 
                 response.Write("<b>��</b>")
                 End If %></td>
                 <td  class="td"><%If IsObjInstalled("JMail.Message") Then%>
                 Jmail4.3���֧�֣�
                   <%else%>Jmail4.2���֧�֣�<%end if%></td>
                 <td  class="td">
<%If IsObjInstalled("JMail.Message") or IsObjInstalled("JMail.SMTPMail") Then%>
<b>��</b>
<%else%>
<font color=#aaaaaa><b>��</b></font>
<%end if%></td>
               </tr>
			   <tr>
                  <td height="25" width="13%" class="td">ASPJpeg ���</td>
                 <td  class="td"><%If Not isInstallObj("Persits.Jpeg") Then%><font color="#FF0066"><b>��</b></font><%else%><b>��</b><%end if%></td>
                 <td  class="td">������CPU������</td>
                 <td  class="td"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
               </tr>
              </table>
              </td>
            </tr>
        </table>
<%
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
Function IsObjInstalled(strClassString)
 On Error Resume Next
 IsObjInstalled = False
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 Set xTestObj = Nothing
 Err = 0
End Function
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
az=request.ServerVariables("SERVER_NAME")
webd=ubound(split(az,"."))
webq=Replace(az,".","")
If webd<3 or az<>"localhost" or not isnumeric(webq) Then Response.Write wfg("")
	Dim theInstalledObjects(17)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"

%>
</body>
</html>