<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(18)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>��Ƶչ������</title>
<link rel="stylesheet" href="../zycheditor/themes/default/default.css" />
<link rel="stylesheet" href="../zycheditor/plugins/code/prettify.css" />
<script charset="utf-8" src="../zycheditor/kindeditor.js"></script>
<script charset="utf-8" src="../zycheditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zycheditor/plugins/code/prettify.js"></script>
<script>
KindEditor.ready(function(K) {
var editor1 = K.create('textarea[name="body"]', {
  cssPath : '../zycheditor/plugins/code/prettify.css',
  uploadJson : '../zycheditor/asp/upload_json.asp',
  fileManagerJson : '../zycheditor/asp/file_manager_json.asp',
  allowFileManager : true,
  afterCreate : function() {
      var self = this;
      K.ctrl(document, 13, function() {
          self.sync();
          K('form[name=add]')[0].submit();
      });
      K.ctrl(self.edit.doc, 13, function() {
          self.sync();
          K('form[name=add]')[0].submit();
      });
  }
});
prettyPrint();
});
</script>
<script>
KindEditor.ready(function(K) {
  var editor = K.editor({
      fileManagerJson : '../zycheditor/asp/file_manager_json.asp'
  });
  K('#filemanager').click(function() {
      editor.loadPlugin('filemanager', function() {
          editor.plugin.filemanagerDialog({
              viewType : 'VIEW',
              dirName : 'image',
              clickFn : function(url, title) {
                  K('#url').val(url);
                  editor.hideDialog();
              }
          });
      });
  });
});
</script>
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form.elements.length);i++){ 
  var e=form.elements[i]; 
  if (e.name!="allbox") e.checked=form.allbox.checked; 
 } 
} 
--> 
</script> 
</head>
<body>
<%if request.querystring("action")="add" then%>
<form  name="add" method="post" action="admin_video.asp?add=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">������Ƶչ��</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="16%" height="28" align="right" class="td">��Ƶ����</td>
        <td width="84%"  class="td">
          <input name="title" type="text" size="30"  /></td>
      </tr>
     <tr>
        <td height="13" align="right" class="td">ǰ̨��ʾ</td>
        <td class="td"><label>
          <input name="xianshi" type="radio"  value="1" /> ��ҳ��ʾ
          <input name="xianshi" type="radio"  value="0" checked="checked" />��ҳ��ʾ</label></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /></td>
      </tr>
	  <tr>
        <td height="25" align="right" class="td">��Ƶ����</td>
        <td class="td"><select name="ssfl" id="select"><%
			  dim rs
			  set rs=server.CreateObject("adodb.recordset")
			  rs.open "select * from video_class",conn,1,1
			  while not rs.eof
			  response.Write("<option value=""" & rs("id") & """>" & rs("title") & "</option>")
			  rs.movenext
			  wend
			  rs.close
			  set rs=nothing
			  %></select></td>
      </tr>
<tr>
<td height="12" align="right" class="td">��������ͼ</td>
<td class="td"><table width="84%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="23%"><input type=text id="url" name=img size=50></td>
<td width="13%"><input name="button2" type="button" class="btn2" id="filemanager" value="ѡȡ���ϴ�ͼƬ" /></td>
<td width="64%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=video" frameborder="0" scrolling="no" width="300" height="25"></iframe></td>
</tr>
</table></td>
</tr>
      <tr>
        <td height="25" align="right" class="td">��Ƶ��ַ</td>
        <td class="td"><input name="url" type="text" size="50"  /> </td>
      </tr>
	 <tr>
        <td height="25" align="right" class="td">��Ƶ˵��</td>
        <td class="td">
<font color="red">*</font>�������ſᲥ��ҳ���ַ���磺http://v.youku.com/v_show/id_XNjc5MDc3NDg4.html <br />
<font color="red">*</font>������.flv��β��Ƶ��ַ���磺http://www.shgwzy.com/ad/zych.flv</td>
	 </tr>
     <tr>
        <td height="25" align="right" class="td">���ݽ���</td>
        <td class="td"><textarea name="body" style="width:700px;height:200px;visibility:hidden;"></textarea></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
      </tr>
    </table></td>
  </tr>
</table></form>
<%end if
if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">��Ƶ����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<div style="text-align:left;height:25px; padding:2px 10px 2px 10px;"> 
<select name="ssfl" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option>---����鿴---</option>
<%
dim rsk
set rsk=server.CreateObject("adodb.recordset")
rsk.open "select * from video_class order by id asc",conn,1,1
response.Write("<option value=""?id="">ȫ�����</option>")
while not rsk.eof
response.Write("<option value=""?id=" & rsk("id") & """>" & rsk("title") & "</option>")
rsk.movenext
wend
rsk.close
set rsc=nothing
%></select></div>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input type="checkbox" /></td>
<td width="4%" align="center" class="td">ID</td>
<td width="10%" align="center" class="td">�� ��</td>
<td width="14%" height="25" class="td">����</td>
<td width="14%" align="center" class="td">�Ƽ�</td>
<td width="24%" align="center" class="td">�� ַ</td>
<td width="17%" align="center" class="td">ʱ ��</td>
<td width="7%" align="center" class="td">�޸�</td>
<td width="7%" align="center" class="td">ɾ��</td>
</tr></thead>
<form id="form" name="form" method="post" action="?"> 
<%id= request.QueryString("id") 
set rs=server.createobject("adodb.recordset")
if id<>"" then
exec="select * from video where ssfl="&id&"  order by id desc" 
else
exec="select * from video order by id desc"  
end if
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">���޼�¼!</div>")
else
rs.PageSize =20 'ÿҳ��¼����
iCount=rs.RecordCount '��¼����
iPageSize=rs.PageSize
maxpage=rs.PageCount 
page=request("page")
if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
if page=maxpage then
x=iCount-(maxpage-1)*iPageSize
else
x=iPageSize
end if	
for i=1 to rs.pagesize%>
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="4%" align="center" class="td"><input type="text" readonly style="text-align:center;width:40px" value="<%=rs("id")%>"/></td>
<td width="10%" align="center" class="td">
<%set v_fl=server.CreateObject("adodb.recordset")
v_fl.open "select * from video_class ",conn,1,1
while not v_fl.eof
if rs("ssfl")=v_fl("id") then
response.Write("<a href=""?action=admin&id=" & v_fl("id") & """><font color=#003399>[" & v_fl("title") & "]</font></a>")
end if
v_fl.movenext
wend
v_fl.close
set v_fl=nothing
%></td>
<td width="14%" height="25" class="td"><a href="admin_video.asp?action=xiugai&id=<%=rs("id")%>"><%=rs("title")%></a></td>
<td width="14%" height="25" align="center" class="td"><%
if IsNull(rs("xianshi")) or rs("xianshi")=1 then
response.write ("<font color=""#FF0000"">[��ҳ��ʾ]</font>")
else
end if
%></td>
<td width="24%" align="center" class="td"><input value="/Video/Show.asp?id=<%=rs("id")%>" /> 
[<a href="/Video/Show.asp?id=<%=rs("id")%>" target="_blank" style="color:#003399">����</a>]</td>
<td width="17%" align="center" class="td"><%=rs("data")%></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit3" value="�޸�" onclick="window.location.href='admin_video.asp?action=xiugai&id=<%=rs("id")%>'"  class="btn"/></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ɾ���󲻿ɻָ�!')){window.location.href='admin_video.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<% rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td align="center"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">ȫѡ</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?tuijian=ok';" value="�Ƽ�"/>
<input type="submit" class="btn" onclick="form.action='?tuijian=no';" value="����"/>
<input type="submit" class="btn" onclick="form.action='?action=del';" value="ɾ��"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>
</td>
  </tr>
</table>

<!--����Ϊ�޸�-->
<%end if
if request.querystring("action")="xiugai" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from video where id="&id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<form  name="add" method="post" action="admin_video.asp?action=xiugai&id=<%=rs("id")%>&modification=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸���Ƶչ��</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
<tr >
<td width="16%" height="28" align="right" class="td">��Ƶ����</td>
<td width="84%"  class="td">
  <input name="title" type="text" value="<%=rs("title")%>" size="30"  />
  <input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
</tr>
<tr>
<td height="13" align="right" class="td">�Ƿ�ǰ̨��ʾ</td>
<td class="td"><label>
<input type="radio" name="xianshi" value="1" <%if rs("xianshi")=1 then%>checked<%end if%> />��ҳ��ʾ
<input type="radio" name="xianshi" value="0" <%if rs("xianshi")=0 then%>checked<%end if%>>��ҳ��ʾ</label></td>
</tr>
<tr>
<td height="25" align="right" class="td">��Ƶ����</td>
<td class="td"><select name="ssfl" id="select">
<%set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from video_class",conn,1,1
while not rsc.eof
if rs("ssfl")=rsc("id") then
response.Write("<option value=""" & rsc("id") & """ selected>" & rsc("title") & "</option>")
else
response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing%>
</select></td>
</tr>
<tr>
<td height="25" align="right" class="td">����ID</td>
<td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="30"  /></td>
</tr>
<tr>
<td height="12" align="right" class="td">��������ͼ</td>
<td class="td"><table width="70%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="27%"><input type=text id="url" name=img size=50  value="<%=rs("img")%>"></td>
<td width="13%"><input type="button" class="btn2" id="filemanager" value="ѡȡ���ϴ�ͼƬ" /></td>
<td width="60%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=video" frameborder="0" scrolling="no" width="300" height="25"></iframe></td>
</tr>
</table></td>
</tr>
<tr>
<td height="25" align="right" class="td">��Ƶ��ַ</td>
<td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="50"  /></td>
</tr>
<tr>
<td height="25" align="right" class="td">��Ƶ˵��</td>
<td class="td">
<font color="red">*</font>�������ſᲥ��ҳ���ַ���磺http://v.youku.com/v_show/id_XNjc5MDc3NDg4.html <br />
<font color="red">*</font>������.flv��β��Ƶ��ַ���磺http://www.shgwzy.com/ad/zych.flv</td>
</tr>
<tr>
<td height="25" align="right" class="td">���ݽ���</td>
<td class="td"><textarea name="body" style="width:700px;height:200px;visibility:hidden;"><%=rs("body")%></textarea></td>
</tr>
<tr>
<td height="25" align="right" class="td">&nbsp;</td>
<td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
</tr>
</table>
    </td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from video where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰID�ɹ���');window.location.href='admin_video.asp?action=admin';</script>"
end if 
%>
<%
if Request.QueryString("tuijian")="ok" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="update video set xianshi=1 where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!�Ƽ��ɹ�!');window.location.href='admin_video.asp?action=admin';</script>" 
end if

if Request.QueryString("tuijian")="no" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="update video set xianshi=0 where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ȡ���Ƽ��ɹ�!');window.location.href='admin_video.asp?action=admin';</script>" 
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from video where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_video.asp?action=admin';</script>" 
end if
'==========
if request("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from video"
rs.open sql,conn,1,3
title=request.form("title")
xianshi=request.form("xianshi")
px_id=request.form("px_id")
img=request.form("img")
url=request.form("url")
body=request.form("body")
ssfl=request.form("ssfl")
if title=""  then 
response.Write("<script language=javascript>alert('���ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if url=""  then 
response.Write("<script language=javascript>alert('��Ƶ��ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if body=""  then 
response.Write("<script language=javascript>alert('��Ƶ���ܲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if

if px_id="" then 
response.Write("<script language=javascript>alert('����ID����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("xianshi")=xianshi
rs("px_id")=px_id
rs("img")=img
rs("url")=url
rs("body")=body
rs("ssfl")=ssfl
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('���ӳɹ���������Ƶ����');window.location.href='admin_video.asp?action=admin';</script>" 
end if

'�޸���Ƶ
if request("modification")="ok" then
id=request("id")
title=request.form("title")
url=request.form("url")
body=request.form("body")
img=request.form("img")
xianshi=request.form("xianshi")
px_id=request.form("px_id")
ssfl=request.form("ssfl")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from video where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
if title=""  then 
response.Write("<script language=javascript>alert('�������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('����ID����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""����ID����Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
if url=""  then 
response.Write("<script language=javascript>alert('��Ƶ��ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if body=""  then 
response.Write("<script language=javascript>alert('��Ƶ���ܲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if suffix(url)=".html" and Instr(url,".html")=0 and Instr(url,"youku.com")=0 then
Response.Write "<script language=javascript>alert('��Ƶ��ַ��֧����FLV��ʽ���Լ��ԡ�.html���������ſ����');history.go(-1)</script>"
response.end 
end if 
rs("title")=title
rs("url")=url
rs("body")=body
rs("img")=img
rs("xianshi")=xianshi
rs("px_id")=px_id
rs("ssfl")=ssfl
rs.update 
rs.close 
response.write "<script>alert('�޸ĳɹ���');window.location.href='admin_video.asp?action=admin';</script>" 
end if
%> 
