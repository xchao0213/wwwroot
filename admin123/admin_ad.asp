<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(25)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>������</title>
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
 
function oCopy(obj){ 
obj.select(); 
js=obj.createTextRange(); 
js.execCommand("Copy") 
alert('���ô����Ѹ��������壡��ճ����ǰ̨��Ҫ��ʾ���ĵط�����!')
} 

<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script> 
</head>
<body>
<%if request.querystring("action")="add" then%>
<form  name="myform" method="post" action="admin_ad.asp?add=ok">
<%id=Request.QueryString("id")%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">���ӹ��</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
<%if id=3 then%>
<tr >
<td width="16%" height="28" align="right" class="td">������� <font color="#FF0000">*</font></td>
<td width="84%"  class="td">
<input name="title" type="text" size="40"  /></td>
</tr>
<tr>
<td width="16%" height="25" align="right" class="td">������� <font color="#FF0000">*</font></td>
<td class="td">     

<select name="lxss" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option value="admin_ad.asp?action=add&id=1" <%if id=1 then%>selected="selected"<%end if%>>ͼƬ���</option>
<option value="admin_ad.asp?action=add&id=2" <%if id=2 then%>selected="selected"<%end if%>>FLASH����</option>
<option value="admin_ad.asp?action=add&id=3" <%if id=3 then%>selected="selected"<%end if%>>���ֹ��</option>
</select>
<label>
<input name="lx" type="hidden" id="lx" value="<%= request.QueryString("id")%> " />
</label></td>
</tr>
      <tr >
<td width="16%" height="28" align="right" class="td">�������� <font color="#FF0000">*</font></td>
<td width="84%"  class="td">
<textarea name="code" cols="40" rows="10"></textarea></td>
</tr>
<%else%>
<tr >
<td width="16%" height="28" align="right" class="td">������� <font color="#FF0000">*</font></td>
<td width="84%"  class="td">
<input name="title" type="text" size="40"  /></td>
</tr>
<tr>
<td width="16%" height="25" align="right" class="td">������� <font color="#FF0000">*</font></td>
<td class="td">     

<select name="lxss" id="jumpMenu" onChange="window.open(this.options[this.selectedIndex].value,'_self')">
<option value="admin_ad.asp?action=add&id=1" <%if id=1 then%>selected="selected"<%end if%>>ͼƬ���</option>
<option value="admin_ad.asp?action=add&id=2" <%if id=2 then%>selected="selected"<%end if%>>FLASH����</option>
<option value="admin_ad.asp?action=add&id=3" <%if id=3 then%>selected="selected"<%end if%>>���ֹ��</option>
</select>
<input name="lx" type="hidden" id="lx" value="<%= request.QueryString("id")%> " /></td>
</tr>

<tr bgcolor="#ffffff">
<td width="16%" height="25" align="right" class="td">ͼƬ��ַ <font color="#FF0000">*</font></td>
<td colspan="2" class="td"><table width="70%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="35%"><input type="text" name="img" size="40" /></td>
<td width="65%"><iframe src="up.asp?formname=myform&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="350" height="21"></iframe></td>
</tr>
</table>
</td>
</tr>
<tr>
<td height="13" align="right" class="td">���ӵ�ַ <font color="#FF0000">*</font></td>
<td class="td"><input name="url" type="text" value="http://" size="40"  /> <label>
�򿪷�ʽ��
  <select name="openfs" id="openfs">
<option value="_blank">_blank</option>
<option value="_parent">_parent</option>
<option value="_self">_self</option>
<option value="_top">_top</option>
</select>
</label></td>
</tr>
<tr>
<td height="6" align="right" class="td">���ߴ�</td>
<td class="td"><input name="width" type="text" size="10"  />
<font color="#FF0000">*</font>
��
<input name="height" type="text" size="10"  />
<font color="#FF0000">*</font>          ��</td>
</tr>
<tr>
<td height="6" align="right" class="td">����˵��</td>
<td class="td"><input name="sm" type="text" size="40"  /></td>
</tr>
<tr>
<td height="25" align="right" class="td"><p>��ע</p>
</td>
<td class="td"><textarea name="bz" cols="50" rows="5"></textarea></td>
</tr><%end if%>
<tr>
<td height="25" align="right" class="td">&nbsp;</td>
<td class="td"><input type="submit" name="button" id="button" value="ȷ���ύ"  class="btn"/></td>
</tr>
</table>
</td>
  </tr>
</table></form>
<%end if
if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="3%" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="3%" align="center" class="td">ID</td>
<td width="26%" height="25" align="center" class="td">�������</td>
<td width="9%" align="center" class="td">����ͼ</td>
<td width="28%" align="center" class="td">���ô���</td>
<td width="8%" align="center" class="td">���</td>
<td width="8%" align="center" class="td">����</td>
<td width="7%" align="center" class="td">�޸�</td>
<td width="8%" align="center" class="td">ɾ��</td>
</tr></thead>
<form id="form1" name="form1" method="post" action="?">
<%set rs=server.createobject("adodb.recordset") 
exec="select * from ad order by id asc" 
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
for i=1 to rs.pagesize %>
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="3%" align="center" class="td"><input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="26%" height="25" align="center" class="td"><a href="admin_ad.asp?action=xiugai&id=<%=rs("id")%>" style="color:#003399"><%=rs("title")%></a> </td>
<td width="9%" align="center" class="td"><%
if IsNull(rs("img")) or rs("img")="" then
else
response.write ("<a href="&rs("img")&"><img src="&rs("img")&" width=""70"" height=""25"" alt=""ͼƬ����ͼ""/></a>")
end if%>
</td>
<td width="28%" align="center" class="td"><input onclick="oCopy(this)" value="&lt;script src=/js/zychad<%=rs("id")%>.js&gt;&lt;/script&gt;" size="30"/></td>
<td width="8%" align="center" class="td"><%=rs("hit")%></td>
<td width="8%" align="center" class="td"><input type="button" name="Submit2" value="����JS" onclick="window.location.href='admin_ad.asp?action=admin&id=<%=rs("id")%>&Generation=ok' "  class="btn"/>
</td>
<td width="7%" align="center" class="td"><input type="button" name="Submit3" value="�޸�" onclick="window.location.href='admin_ad.asp?action=xiugai&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="8%" align="center" class="td">
<input type="button" name="Submit" value="ɾ��" onclick="javascript:if(confirm('ȷ��ɾ����ǰ��棿ɾ���󲻿ɻָ�!')){window.location.href='admin_ad.asp?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td width="3%" height="30" align="center" class="td"><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td width="3%" height="30" align="center" class="td">ȫѡ</td>
<td colspan="7" class="td"><input type="submit" class="btn" onclick="form.action='?action=admin&act=del';" value="ɾ��"/>
<%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr></form>
<tr><td colspan="9">&nbsp;ע�⣺���ӹ����޸Ĺ����������JS�ٵ��ã�</td></tr>
</table></td>
  </tr>
</table>
<!--����Ϊ�޸�-->
<%end if
if request.querystring("action")="xiugai" then
exec="select * from ad where id="& request.QueryString("id") 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
id=rs("id")
lx=rs("lx")
%>
<form  name="add" method="post" action="admin_ad.asp?id=<%=rs("id")%>&xiugai=ok">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�޸Ĺ��</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
          <tr >
        <td width="16%" height="14" align="right" class="td"><font color="#FF0000">*</font>�������</td>
        <td width="84%" height="25"  class="td"><strong>
		<%if lx=1 then
		response.Write("ͼƬ���")
		elseif lx=2 then
		response.Write("FLASH���")
		else
		response.Write("���ֹ��")
		end if
		%></strong>
          <input name="lx" type="hidden" id="lx" value="<%=rs("lx")%>" /></td>
          </tr>
       <%if lx=3 then%>  
      <tr >
        <td width="16%" height="28" align="right" class="td"><font color="#FF0000">*</font>�������</td>
        <td width="84%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="40"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="40"  /></td>
      </tr> 
<tr >
                          <td height="14" align="right" class="td"><font color="#FF0000">*</font>��������</td>
<td width="84%"  class="td"><textarea name="code" cols="40" rows="10"><%=rs("code")%></textarea></td>
                        </tr><%else%>
         <tr >
        <td width="16%" height="28" align="right" class="td"><font color="#FF0000">*</font>�������</td>
        <td width="84%"  class="td">
          <input name="title" type="text" value="<%=rs("title")%>" size="40"  />
          <input name="id" type="hidden" value="<%=rs("id")%>" size="40"  /></td>
      </tr> 
      <tr bgcolor="#ffffff">
        <td width="16%" height="25" align="right" class="td"><font color="#FF0000">*</font>ͼƬ��ַ</td>
        <td colspan="2" class="td"><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="32%"><input type="text" name="img" id="url" size="40" value="<%=rs("img")%>"></td>
    <td width="9%"><input name="button2" type="button" id="filemanager" value="ѡȡ���ϴ�" /></td>
    <td width="59%"><iframe src="up.asp?formname=add&inputname=img&uploadstyle=Ad" frameborder="0" scrolling="no" width="350" height="21"></iframe></td>
  </tr>
</table> 
</td>
  </tr>
    <tr>
        <td height="13" align="right" class="td"><font color="#FF0000">*</font>���ӵ�ַ</td>
        <td class="td"><input name="url" type="text" value="<%=rs("url")%>" size="40"  />
          �򿪷�ʽ��
          <select name="openfs" id="openfs">
            <option value="_blank"  <%if rs("openfs")="_blank" then%>selected="selected"<%end if%>>_blank</option>
            <option value="_parent" <%if rs("openfs")="_parent" then%>selected="selected"<%end if%>>_parent</option>
            <option value="_self" <%if rs("openfs")="_self" then%>selected="selected"<%end if%>>_self</option>
            <option value="_top" <%if rs("openfs")="_top" then%>selected="selected"<%end if%>>_top</option>
          </select></td>
    </tr>
    <tr>
      <td height="6" align="right" class="td">���ߴ�</td>
      <td class="td"><input name="width" type="text" value="<%=rs("width")%>" size="10"  />
        <font color="#FF0000">*</font>
        ��
          <input name="height" type="text" value="<%=rs("height")%>" size="10"  />
          <font color="#FF0000">*</font>          ��</td>
    </tr>
    <tr>
      <td height="6" align="right" class="td">����˵��</td>
      <td class="td"><input name="sm" type="text" value="<%=rs("sm")%>" size="40"  /></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td"><p>��ע</p>          </td>
        <td class="td"><textarea name="bz" cols="50" rows="5"><%=rs("bz")%></textarea></td>
      </tr>
    <tr>
        <td height="13" align="right" class="td">����ʱ��</td>
        <td class="td"><input name="data" type="text" value="<%=rs("data")%>" size="40"   onfocus="show_cele_date(data,'','',data)"/></td>
      </tr>  <%end if%>
      <tr>
      <td height="12" align="right" class="td">&nbsp;</td>
      <td class="td"><input type="submit" name="button" id="button" value="�ύ�޸�"  class="btn"/></td>
    </tr>
  
    </table>
</td>
  </tr>
</table></form>
<%end if%>
</body>
</html>
<%
if Request.QueryString("add")="ok" then
set rs=server.createobject("adodb.recordset")
sql="select * from ad"
rs.open sql,conn,1,3
title=request.form("title")
url=request.form("url")
width=request.form("width")
height=request.form("height")
img=request.form("img")
lx=request.form("lx")
sm=request.form("sm")
bz=request.form("bz")
openfs=request.form("openfs")
code=request.form("code")
if title=""  then 
response.Write("<script language=javascript>alert('������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if lx=1 or lx=2  then
if img=""  then 
response.Write("<script language=javascript>alert('���ͼƬ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if url=""  then 
response.Write("<script language=javascript>alert('������ӵ�ַ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if width=""  then 
response.Write("<script language=javascript>alert('����������!');history.go(-1)</script>") 
response.end 
end if
if height=""  then 
response.Write("<script language=javascript>alert('��������߶�!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("width")) or not isNumeric(request("height"))then
response.write("<script>alert(""���ߴ���붼Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
else
if code=""  then 
response.Write("<script language=javascript>alert('���ֹ�����ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
end if
rs.addnew
rs("title")=title
rs("url")=url
rs("width")=width
rs("height")=height
rs("img")=img
rs("lx")=lx
rs("sm")=sm
rs("bz")=bz
rs("code")=code
rs("openfs")=openfs
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('������ӳɹ��˹���');window.location.href='admin_ad.asp?action=admin';</script>" 
end if
%>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from ad where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰ���ɾ���ɹ���');window.location.href='admin_ad.asp?action=admin';</script>"
end if
if Request.QueryString("act")="del" then
if Request("id")="" then
Response.Write "<script>alert('����!��ѡ��Ҫ�����ļ�¼!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [ad] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('��ϲ!ɾ���ɹ�!');window.location.href='admin_ad.asp?action=admin';</script>" 
end if

if Request.QueryString("xiugai")="ok" then
id=request("id")
title=request.form("title")
lx=request.form("lx")
img=request.form("img")
url=request.form("url")
width=request.form("width")
height=request.form("height")
sm=request.form("sm")
bz=request.form("bz")
code=request.form("code")
data=request.form("data")
openfs=request.form("openfs")
	if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
	Response.End()
	end if
	SQL="Select * from ad where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
	Response.End()
	end if
	if title=""  then 
if lx=1 or lx=2  then
response.Write("<script language=javascript>alert('������Ʋ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
if img=""  then 
response.Write("<script language=javascript>alert('���ͼƬ����Ϊ��!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("width")) or not isNumeric(request("height"))then
response.write("<script>alert(""���ߴ���붼Ϊ���֣�""); history.go(-1);</script>")
response.end
end if
elseif lx=3 then
if code=""  then 
response.Write("<script language=javascript>alert('�������ݲ���Ϊ��!');history.go(-1)</script>") 
response.end 
end if
else
end if
rs("title")=title
rs("url")=url
rs("width")=width
rs("height")=height
rs("img")=img
rs("lx")=lx
rs("sm")=sm
rs("bz")=bz
rs("openfs")=openfs
rs("data")=data
rs("code")=code
rs.update 
rs.close 
set js = server.CreateObject("ADODB.RecordSet")
sql="select * from ad where id="&request.QueryString("id")
set js = conn.Execute (Sql) 
if js("lx")=1 then
goaler = goaler + "<a href="""&dir&"Include/Showbody.asp?action=adurl&id="& js("id")&""" target="""& js("openfs")&"""><img src="""& js("img")&""" width="""& js("width")&""" height="""& js("height")&"""  title="""& js("sm")&"""></a>"  
elseif js("lx")=2 then
goaler = goaler + "<embed src="""&js("img")&""" quality=""height"" type=""application/x-shockwave-flash""  width="""& js("width")&""" height="""& js("height")&""" ></embed>" 
else
goaler = goaler + ""& js("code")&"" 
end if
'����JS�ļ�
goaler = "" + goaler + ""
goaler = "document.write('" & goaler & "')"
FolderPath = Server.MapPath("../")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.CreateTextFile(FolderPath&"/js/zychad"&js("id")&".js")
fout.WriteLine goaler
'�ر�����
fout.close
set fout = nothing
js.close
set js = nothing
conn.close
set conn=nothing
response.write "<script>alert('����޸ĳɹ���');window.location.href='admin_ad.asp?action=admin';</script>" 
end if
%> 
<%
if Request.QueryString("Generation")="ok" then
set js = server.CreateObject("ADODB.RecordSet")
sql="select * from ad where id="&request.QueryString("id")
set js = conn.Execute (Sql) 
if js("lx")=1 then
goaler = goaler + "<a href="""&dir&"Include/Showbody.asp?action=adurl&id="& js("id")&""" target="""& js("openfs")&"""><img src="""& js("img")&""" width="""& js("width")&""" height="""& js("height")&"""  title="""& js("sm")&"""></a>"  
elseif js("lx")=2 then
goaler = goaler + "<embed src="""&js("img")&""" quality=""height"" type=""application/x-shockwave-flash""  width="""& js("width")&""" height="""& js("height")&""" ></embed>" 
else
goaler = goaler + ""& js("code")&"" 
end if
'����JS�ļ�
goaler = "" + goaler + ""
goaler = "document.write('" & goaler & "')"
FolderPath = Server.MapPath("../")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.CreateTextFile(FolderPath&"/js/zychad"&js("id")&".js")
fout.WriteLine goaler
'�ر�����
fout.close
set fout = nothing
js.close
set js = nothing
conn.close
set conn=nothing
Response.Write "<script>alert('���JS�Ѿ�����!');window.location.href='admin_ad.asp?action=admin';</script>" 
end if
%>