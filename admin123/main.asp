<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%az=""&request.ServerVariables("SERVER_NAME")&""
wz=request.ServerVariables("HTTP_HOST")&request.ServerVariables("URL")
wurl=right(left(wz,instrrev(wz,"/")),(len(left(wz,instrrev(wz,"/")))-len(request.ServerVariables("HTTP_HOST")))-1)%>  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>服务器信息</title>
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
<div class="box"><a href="javascript:void(0)" onclick ="document.getElementById('pop').style.display='none'" style="position:absolute;right:0px;color:#000; margin:2px 10px;" title="关闭窗口" id="closeBtn">×关闭</a>
<p class="t1">您的网站后台默认目录为“<a>admin</a>”,为了安全起见请修改为其它名称</p>
<form id="form" name="form" method="post" action="?admin=ok">
<p style="text-align:center"><input name="admin" type="text" class="value" id="admin" value="" placeholder="您的后台名称"/><input class="button" name="button" type="submit" value="确认修改" /></p>
</form>
<p class="t2">修改完成后请重新打开后台！</p>
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
response.Write("<script language=javascript>alert('目录名称不能为空!');history.go(-1)</script>") 
response.end 
end if
set fso=Server.CreateObject("Scripting.FileSystemObject") 
set file=fso.GetFolder(Server.MapPath("../admin")) 
file.name=request.form("admin")
set fso=nothing
Response.Write("<div id=""pop"" style=""display:block;""><div class=""box"">")
Response.Write("<h1>修改成功！</h1>")
Response.Write("<p style=""font-size:16px;color:#F00;font-family:'微软雅黑'"">您的后台登陆目录已经修改为“"&newadmin&"”,新牢记</p>")
Response.Write("<p><a style=""font-size:20px;text-align:center;"" href=""/"&newadmin&""" target=""_blank"">[您可以点击这里打开新的后台]</a></p>")
Response.Write("</div></div>")
end if
End If
if Instr(request.ServerVariables("SERVER_SOFTWARE"),"NetBox")>0 then%>
<div class="pop" id="pip" style="display:none;">
<div class="box"><a href="javascript:void(0)" onclick ="document.getElementById('pip').style.display='none'" style="position:absolute;right:0px;color:#000; margin:2px 10px;" title="关闭窗口" id="closeBtn">×关闭</a>
<p class="t2">在测试本网站时，首先要确认网站是否在<br />
<a style="color:#F00">本地Windows-IIS环境下 或 上传到FTP空间进行测试</a></p>
<p>而不是下面这类小工具，这类小工具会影响网站功能使用，如后台不能上传图片等</p>
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
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">站点信息</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
<tr bgcolor="#F1F5F8" >
    <td height="25" colspan="6" class="td">
  <span class="mian_bt">感谢您使用高屋置业企业网站管理系统<%=zychversion%>版，许可证号：<%=authorized%> DATA:<%=data%></span>
  <div class="mian_body"><ul style="padding-left:25px; margin:0px">
  <li>使用本系统注意事项：</li>
  <li>本程序由高屋置业开发，版权归高屋置业所有！</li>
  <li>本系统为商业程序，未经授权严禁使用，如有需要请于本公司联系，请尊重作者劳动成果，任何单位、个人不得修改版权和进行二次销售！</li>
  <li>为了后期能更好的解决您网站使用过程中遇到的问题，请不要删除如：<%=data%>此类的版本号,以便日后便于识别!</li>
  <li>您在使用过程，如发现BUG！请及时联系高屋置业程序员（QQ：495573637 高屋置业人），本公司将特别感谢您对我们的支持！www.shgwzy.com</li>
  </ul>
  </div>	
  
    </td>
</tr>
<tr >
      <td width="10%" height="25" align="right" class="td">管理员个数：</td>
<td width="20%"  class="td"><b><%=conn.execute("select count(*) from admin")(0)%> </b>个</td>
      <td width="8%" align="right"  class="td">产品个数：</td>
      <td width="18%"  class="td"><b><%=conn.execute("select count(*) from products")(0)%></b> 个</td>
      <td width="12%" align="right"  class="td">新订单/全部订单：</td>
      <td width="32%"  class="td"><b><font color="#FF0000"><%=conn.execute("select count(*) from orders where state=1")(0)%></font>/<%=conn.execute("select count(*) from orders ")(0)%></b> [<a href="admin_orders.asp?action=admin">查看</a>]</td>
</tr>
    <tr>
      <td width="10%" height="25" align="right" class="td">新闻条数：</td>
      <td class="td"><b><%=conn.execute("select count(*) from news")(0)%></b> 条</td>
      
      <td align="right" class="td">收到的简历：</td>
      <td class="td"><b><%=conn.execute("select count(*) from Resume")(0)%></b> 份</td>
      <td align="right" class="td">新留言/全部留言：</td>
      <td class="td"><b><font color="#FF0000"><%=conn.execute("select count(*) from ly where hf is null")(0)%></font>/<%=conn.execute("select count(*) from ly ")(0)%></b> [<a href="admin_message.asp?action=admin">查看</a>]</td>
    </tr>
    <tr>
      <td width="10%" height="25" align="right" class="td">下载个数：</td>
      <td class="td"><b><%=conn.execute("select count(*) from download")(0)%></b> 个</td>
      <td align="right" class="td">会员个数：</td>
      <td align="left" class="td"><b><%=conn.execute("select count(*) from [user]")(0)%></b> 个</td>
      <td align="right" class="td">照片张数：</td>
      <td align="left" class="td"><b><%=conn.execute("select count(*) from photo")(0)%></b> 张</td>
    </tr>
  <tr>
    <td height="25" align="right" class="td">&nbsp;</td>
    <td class="td">&nbsp;</td>
    <td align="right" class="td">团队人数：</td>
    <td class="td"><b><%=conn.execute("select count(*) from tuandui")(0)%></b> 人</td>
    <td align="right" class="td">视频个数：</td>
    <td class="td"><b><%=conn.execute("select count(*) from [video]")(0)%></b> 个</td>
  </tr>
</table></td>
</tr>
</table></div>
       </td>
	  </tr>
</table>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
            <tr>
              <td height="28" background="images/bg_list.gif" bgcolor="#F7F7F7"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">服务器信息</div></td>
            </tr>
            
            <tr>
              <td bgcolor="#FFFFFF" ><table width="100%" border="0" align="center" cellpadding="4" cellspacing="0" class="stable" >
                
<tr >
                  <td height="25" width="13%" class="td">程序绝对路径:</td>
                <td width="43%"  class="td"><%=Server.MapPath("nowaasp.asp")%></td>
                <td width="10%"  class="td">服务器脚本引擎：</td>
                <td width="34%"  class="td"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
            <tr>
                  <td height="25" width="13%" class="td">站点物理路径：</td>
                <td class="td"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
                <td class="td">路径信息：</td>
                <td class="td"><%=request.ServerVariables("PATH_INFO")%></td>
            </tr>
              
                <tr>
                  <td height="25" width="13%" class="td">请求机器IP地址：</td>
                  <td class="td"><%=request.ServerVariables("REMOTE_ADDR")%></td>
                  <td class="td">服务器内部IP：</td>
                  <td class="td"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
                </tr>
                <tr>
                  <td height="25" width="13%"  class="td">SCRIPT虚拟路径：</td>
                  <td  class="td"><%=request.ServerVariables("SCRIPT_NAME")%></td>
                  <td  class="td">服务器IP地址：</td>
                  <td  class="td"><%=request.ServerVariables("SERVER_NAME")%></td>
                </tr>
              <tr>
                  <td height="25" width="13%" class="td">服务器端口：</td>
                <td class="td"><%=request.ServerVariables("SERVER_PORT")%></td>
                <td class="td">协议名称和版本：</td>
                <td class="td"><%=request.ServerVariables("SERVER_PROTOCOL")%></td>
              </tr>
                 <tr>
                  <td height="25" width="13%" class="td">服务器IIS版本：</td>
                  <td class="td"><%=request.ServerVariables("SERVER_SOFTWARE")%></td>
                  <td class="td">脚本超时时间：</td>
                  <td class="td"><%=Server.ScriptTimeout%>秒</td>
                </tr>
               <tr>
                  <td height="25" width="13%" class="td">服务器操作系統：</td>
                 <td  class="td"><%=Request.ServerVariables("OS")%>&nbsp;</td>
                 <td  class="td">服务器CPU数量：</td>
                 <td  class="td"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
               </tr>
               <tr>
                 <td height="25" class="td">FSO文本读写：</td>
                 <td  class="td">
				 <% if IsObjInstalled("Scripting.FileSystemObject") = False Then
				 response.Write("<font color=#aaaaaa><b>×</b></font>")
                 Else 
                 response.Write("<b>√</b>")
                 End If %></td>
                 <td  class="td"><%If IsObjInstalled("JMail.Message") Then%>
                 Jmail4.3组件支持：
                   <%else%>Jmail4.2组件支持：<%end if%></td>
                 <td  class="td">
<%If IsObjInstalled("JMail.Message") or IsObjInstalled("JMail.SMTPMail") Then%>
<b>√</b>
<%else%>
<font color=#aaaaaa><b>×</b></font>
<%end if%></td>
               </tr>
			   <tr>
                  <td height="25" width="13%" class="td">ASPJpeg 组件</td>
                 <td  class="td"><%If Not isInstallObj("Persits.Jpeg") Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
                 <td  class="td">服务器CPU数量：</td>
                 <td  class="td"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
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