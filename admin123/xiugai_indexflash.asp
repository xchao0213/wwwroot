<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%
if session("qx")=1 then
response.Write ("<div align=""center""><font style=""color:red; font-size:25px;"">��û�й����ģ���Ȩ�ޣ�</font></div>")
response.End
End If
%>
<% 
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from indexflash where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>�޸Ļõ�</title>
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
</head>
<body>
<form  name="myform" method="post" action="xiugai_indexflash.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸Ļõ�</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8" >
        <td height="28" width="16%" class="td">����</td>
        <td  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="30"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="30"  /></td>
      </tr>
      <tr bgcolor="#ffffff">
        <td height="25" width="16%" class="td">ͼƬ��ַ <font color="#FF0000">*</font></td>
        <td colspan="2" class="td"><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%"><input type="text" name="img" id="url" size="30" value="<%=rs("img")%>"></td>
    <td width="11%"><input name="button2" type="button" id="filemanager" value="ѡȡ���ϴ�ͼƬ" class="btn" style="height:20px;" /></td>
    <td width="59%"><iframe src="up.asp?formname=myform&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="340" height="21"></iframe></td>
  </tr>
</table>          </td>
        </tr>
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" width="16%" class="td">���ӵ�ַ</td>
        <td class="td"><input name="link" type="text" value="<%=rs("link")%>" size="30"  /> 
          ����������#��</td>
      </tr>
	  
    <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">����ID</td>
        <td class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="10"  /></td>
      </tr>
	<!--
	 <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" width="16%" class="td">�õƶ�����ʽ</td>
          <td class="td"><select name="slicing">
          <option value=""<%if rs("slicing")="" then%>selected<%end if%>>ˮƽ��ֱ����</option>
          <option value="horizontal" <%if rs("slicing")="horizontal" then%>selected<%end if%>>horizontalˮƽ</option>
          <option value="vertical"  <%if rs("slicing")="vertical" then%>selected<%end if%>>vertical��ֱ</option>
        </select>
          <label>
		<select name="direction">
		 <option value="" <%if rs("direction")="" then%>selected<%end if%>>ѡ����</option>
          <option value="left" <%if rs("direction")="left" then%>selected<%end if%>>left����</option>
          <option value="right"<%if rs("direction")="right" then%>selected<%end if%>>right����</option>
          <option value="down" <%if rs("direction")="down" then%>selected<%end if%>>down����</option>
          <option value="up" <%if rs("direction")="up" then%>selected<%end if%>>up����</option>
        </select>
          </label>
          <label>		
		<select name="num">
          <option value="" <%if rs("num")="" then%>selected<%end if%>>��Ƭ����</option>
          <option value="1" <%if rs("num")="1" then%>selected<%end if%>>1Ƭ</option>
		  <option value="2" <%if rs("num")="2" then%>selected<%end if%>>2Ƭ</option>
          <option value="3" <%if rs("num")="3" then%>selected<%end if%>>3Ƭ</option>
          <option value="4" <%if rs("num")="4" then%>selected<%end if%>>4Ƭ</option>
          <option value="5" <%if rs("num")="5" then%>selected<%end if%>>5Ƭ</option>
		  <option value="6" <%if rs("num")="6" then%>selected<%end if%>>6Ƭ</option>
        </select>
          </label>
          <label>
		<select name="shader">
          <option value="" <%if rs("shader")="" then%>selected<%end if%>>����Ƭ��Ӱ</option>
          <option value="phong" <%if rs("shader")="phong" then%>selected<%end if%>>����Ƭ��Ӱ</option>
        </select>
          </label>
          <label>
		<select name="delay">
          <option value="" <%if rs("delay")="" then%>selected<%end if%>>�ӳ�ʱ��</option>
          <option value="0.03" <%if rs("delay")="0.03" then%>selected<%end if%>>�ӳ�0.03</option>
          <option value="0.05" <%if rs("delay")="0.05" then%>selected<%end if%>>�ӳ�0.05</option>
          <option value="0.1" <%if rs("delay")="0.1" then%>selected<%end if%>>�ӳ�0.1</option>
        </select>
          </label>
            </td>
        </tr>-->
	<tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#FFFFFF'" bgcolor="#FFFFFF">
        <td height="25" class="td">������ɫ <font color="#FF0000">*</font></td>
        <td class="td"><input name="coler" type="text" style="background:<%=rs("coler")%>; color:#FFFFFF" value="<%=rs("coler")%>"  size="10"  /> 
        ��������ɫ</td>
      </tr>	 
      <tr onmouseover="style.backgroundColor='#EEEEEE'" onmouseout="style.backgroundColor='#F1F5F8'" bgcolor="#F1F5F8">
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><input type="submit" name="button" id="button" value="ȷ���޸�"  class="btn"/></td>
      </tr>
    
    </table>
</td>
  </tr>
</table></form>
</body>
</html>
<% 
if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
img=request.form("img")
link=request.form("link")
px_id=request.form("px_id")
direction=request.form("direction")
slicing=request.form("slicing")
num=request.form("num")
shader=request.form("shader")
delay=request.form("delay")
coler=request.form("coler")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from indexflash where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
if title=""  then 
response.Write("<script language=javascript>alert('�����ⲻ��Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if link=""  then 
response.Write("<script language=javascript>alert('���ӵ�ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if img="" then 
response.Write("<script language=javascript>alert('ͼƬ��ַ����Ϊ��!');history.go(-1)</script>") 
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
rs("title")=title
rs("img")=img
rs("link")=link
rs("px_id")=px_id
rs("direction")=direction
rs("slicing")=slicing
rs("num")=num
rs("shader")=shader
rs("delay")=delay
rs("coler")=coler
rs.update 
rs.close 
response.write "<script>alert('��ҳ�õ��޸ĳɹ���');window.location.href='admin_indexflash.asp';</script>" 
end if
%> 
