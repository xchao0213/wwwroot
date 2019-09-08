<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/inc.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(3)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>网站菜单管理</title>
<style>
.menu td { border-bottom:1px #E4E4E4 solid}
</style>
</head>
<body>
<%if request.querystring("menu")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">网站导航管理 
	<input type="button" value="添加主导航"  class="btn2" onClick="window.location.href='?menu=add_onemenu'"/>
	<input type="button"  value="添加二级导航"  class="btn2" onClick="window.location.href='?menu=add_menu_nav'"/></div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
      <tr>
        <td bgcolor="#FFFFFF">
        </td>
      </tr>
      <tr style="background-color:#F1F5F8" class="menu">
        <td colspan="2"><table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr style="font-family:'微软雅黑'">
<td width="54" align="center"  class="td">ID</td>
<td width="175" height="25" align="center">导航标题</td>
<td width="112" height="25">导航英文</td>
<td width="118" height="25" align="center">导航地址</td>
<td width="82" align="center" >所 属</td>
<td width="115" align="center" >打开窗口</td>
<td width="114" align="center">显示位置</td>
<td width="79" align="center">排序</td>
<td width="60" align="center" >修改</td>
<td width="69" align="center">删除</td>
</tr></thead>
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from menu order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无菜单!</div>")
else
while not rs.eof%>       
<form action="admin_menu.asp?menu=admin&xiugai=ok" method="post" name="add">
<tr style="font-family:'微软雅黑'; background:#F7F7F7" onMouseOver="style.backgroundColor='#FFC'" onMouseOut="style.backgroundColor='#F7F7F7'" bgcolor="#F7F7F7">
<td align="center"  class="td"><input name="id" type="hidden" size="15"  value="<%=rs("id")%>"/>  <input type="text" readonly style="text-align:center; width:40px" value="<%=rs("id")%>"/></td>
<td width="175" height="25"><input name="title" type="text" value="<%=rs("title")%>"/></td>
<td width="112" height="25"><input name="en_title" type="text" value="<%=rs("en_title")%>" size="10"/></td>
<td height="25" align="center"><input name="url" type="text" size="15"  value="<%=rs("url")%>"/></td>
<td width="82" align="center" >一级导航</td>
<td width="115" align="center" ><select name="openfs">
<option value="_self" <%if rs("openfs")="_self" then%>selected="selected"<%end if%>>原窗口</option>
<option value="_blank" <%if rs("openfs")="_blank" then%>selected="selected"<%end if%>>新窗口</option>
</select></td>
<td width="114" align="center"><select name="yc">
<option style="color:#FF0000" value="1" <%if rs("yc")=1 then%>selected="selected"<%end if%>>主菜单显示</option>
<option value="2" <%if rs("yc")=2 then%>selected="selected"<%end if%>>顶菜单显示</option>
<option style="color:#0066CC" value="3" <%if rs("yc")=3 then%>selected="selected"<%end if%>>隐藏不显示</option>
</select></td>
<td width="79" align="center"><input name="px_id" type="text" style="text-align:center; width:40px" value="<%=rs("px_id")%>"/></td>
<td width="60" align="center" ><input type="submit" name="button2" id="button2" value="修改"  class="btn"/></td>
<td width="69" align="center"><input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_menu.asp?menu=admin&act=del&id=<%=rs("id")%>';}else{history.go(0);}"  class="btn"/></td>
</tr>
</form>
<% set rsb=server.createobject("adodb.recordset") 
exec="select * from [menu_nav] where ssfl="&rs("id")&" order by px_id asc  " 
rsb.open exec,conn,1,1 
if rsb.eof and rsb.bof then
response.Write("")
end if
do while not rsb.eof%>
<tr>
<form action="admin_menu.asp?menu=admin&xiugai_nav=ok" method="post" name="add">
<td align="right"><input name="id" type="hidden" size="15"  value="<%=rsb("id")%>"/></td>
<td>├─<span style="font-family: 'Arial Black', Gadget, sans-serif"><%=rsb("px_id")%>.</span>
  <input name="title" type="text" size="14"  value="<%=rsb("title")%>"/></td>
<td width="112" height="25">&nbsp;</td>
<td align="center"><input name="url" type="text" size="15"  value="<%=rsb("url")%>"/></td>
<td align="center"><select name="ssfl" id="ssfl">
<%
set rss=server.CreateObject("adodb.recordset")
rss.open "select * from [menu] where yc=1 order by px_id asc",conn,1,1
while not rss.eof
if rsb("ssfl")=rss("id") then
response.Write("<option value=""" & rss("id") & """ selected>" & rss("title") & "</option>")
else
response.Write("<option value=""" & rss("id") & """>" & rss("title") & "</option>")
end if
rss.movenext
wend
rss.close
set rss=nothing
%>
  </select></td>
<td align="center"><select name="openfs">
  <option value="_self" <%if rsb("openfs")="_self" then%>selected="selected"<%end if%>>原窗口</option>
  <option value="_blank" <%if rsb("openfs")="_blank" then%>selected="selected"<%end if%>>新窗口</option>
</select></td>
<td align="center"><select name="yc">
<option style="color:#FF0000" value="1" <%if rsb("yc")=1 then%>selected="selected"<%end if%>>二级显示</option>
<option style="color:#0066CC" value="3" <%if rsb("yc")=3 then%>selected="selected"<%end if%>>隐藏不显示</option>
 </select></td>
<td align="center"> <input name="px_id" type="text" style="text-align:right; width:40px" value="<%=rsb("px_id")%>"/></td>
<td align="center"><input type="submit" name="button2" id="button2" value="修改"  class="btn"/></td>
<td align="center"><input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_menu.asp?menu=admin&act=nav_del&id=<%=rsb("id")%>';}else{history.go(0);}"  class="btn"/></td>
</form>
</tr>
<% rsb.movenext
loop
rsb.close
set rsb=nothing %>
<%rs.movenext
wend
end if
rs.close
set rs=nothing
%>
</table>
</td>
</tr>
    </table></td>
  </tr>
</table>
<% end If %>
<%if request.querystring("menu")="add_onemenu" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?menu=add_onemenu&action=add" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加一级导航 	
	    <input type="button" value="管理导航菜单"  class="btn2" onClick="window.location.href='?menu=admin'"/>
	<input type="button" value="添加二级导航"  class="btn2" onClick="window.location.href='?menu=add_menu_nav'"/></div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="16%" height="25" align="right" class="td">中文名称</td>
        <td width="84%"  class="td"><input name="title" type="text" size="30"  /></td>
      </tr>
      <tr >
        <td width="16%" height="25" align="right" class="td">英文名称</td>
        <td width="84%"  class="td"><input name="etitle" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td width="16%" height="13" align="right" class="td">链接地址</td>
        <td class="td"><input name="url" type="text" size="30"  /></td>
      </tr>
      <tr>
        <td width="16%" height="12" align="right" class="td">打开方式</td>
        <td class="td"><select name="openfs" id="openfs">
          <option value="_self">原窗口</option>
          <option value="_blank">新窗口</option>
        </select>
          * </td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /> 
          数字越小越靠前。</td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">菜单显示位置</td>
        <td class="td"><select name="yc" id="yc">
          <option value="1">主菜单显示</option>
          <option value="2">顶菜单显示</option>
		  <option value="3">隐藏不显示</option>
        </select> 
          隐藏后前台将不显示该菜单。</td>
      </tr>
      <tr>
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><label>
        <input type="submit" name="button" id="button" value="提交数据"  class="btn"/>
        </label></td>
      </tr>
      
    </table></td>
  </tr></form>
</table>

<% end If %>
<%if request.querystring("menu")="add_menu_nav" then%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?action=nav_add" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加二级导航
	    <input type="button" value="管理导航菜单"  class="btn2" onClick="window.location.href='?menu=admin'"/>
	<input type="button" value="添加一级导航"  class="btn2" onClick="window.location.href='?menu=add_onemenu'"/>
	</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="16%" height="25" align="right" class="td">二级导航名称</td>
        <td width="84%"  class="td"><input name="title" type="text" id="title" value="0" size="30"  /></td>
      </tr>
      <tr >
        <td width="16%" height="25" align="right" class="td">所属一级导航</td>
        <td width="84%"  class="td"><select name="ssfl" id="select">
		  <%
          dim rsc
          set rsc=server.CreateObject("adodb.recordset")
          rsc.open "select * from menu where yc=1 order by px_id asc",conn,1,1
          while not rsc.eof
          response.Write("<option value=""" & rsc("id") & """>" & rsc("title") & "</option>")
          rsc.movenext
          wend
          rsc.close
          set rsc=nothing
          %></select>
          
<select id="impower" name="impower" onChange="this.form.url.value=this.value;this.form.title.value=this.options[this.selectedIndex].text;">
<option value="#">=====选择已有分类====</option>
<%=all_navoption%>
</select>
          
          
          </td>
      </tr>
      <tr>
        <td width="16%" height="13" align="right" class="td">链接地址</td>
        <td class="td"><input name="url" type="text" id="url" value="0" size="30"  /></td>
      </tr>
      <tr>
        <td width="16%" height="12" align="right" class="td">打开方式</td>
        <td class="td"><select name="openfs" id="openfs">
          <option value="_self">原窗口</option>
          <option value="_blank">新窗口</option>
        </select>
          * </td>
      </tr>
      <tr>
        <td width="16%" height="25" align="right" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /> 
          数字越小越靠前。</td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">菜单显示位置</td>
        <td class="td"><select name="yc" id="yc">
          <option value="1">显示</option>
		  <option value="3">不显示</option>
        </select> 
          隐藏后前台将不显示该菜单。</td>
      </tr>
      <tr>
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><label>
        <input type="submit" name="button" id="button" value="提交数据"  class="btn"/>
        </label></td>
      </tr>
      
    </table></td>
  </tr></form>
</table>
<% end If %>	
</body>
</html>
<%
if request("action")="add" then
set rs=server.createobject("adodb.recordset")
sql="select * from menu"
rs.open sql,conn,1,3
title=request.form("title")
url=request.form("url")
openfs=request.form("openfs")
px_id=request.form("px_id")
yc=request.form("yc")
if title=""  then 
response.Write("<script language=javascript>alert('导航名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if url=""  then 
response.Write("<script language=javascript>alert('链接地址不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert("" 错误！排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("url")=url
rs("openfs")=openfs
rs("px_id")=px_id
rs("yc")=yc
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('主导航增加成功！');window.location.href='admin_menu.asp';</script>" 
end if
%>
<%
if request("action")="nav_add" then
set rs=server.createobject("adodb.recordset")
sql="select * from menu_nav"
rs.open sql,conn,1,3
title=request.form("title")
ssfl=request.form("ssfl")
url=request.form("url")
openfs=request.form("openfs")
px_id=request.form("px_id")
yc=request.form("yc")
if title=""  then 
response.Write("<script language=javascript>alert('二级导航名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if url=""  then 
response.Write("<script language=javascript>alert('链接地址不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert("" 错误！排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("ssfl")=ssfl
rs("url")=url
rs("openfs")=openfs
rs("px_id")=px_id
rs("yc")=yc
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('菜单增加成功！');window.location.href='admin_menu.asp?menu=admin&';</script>" 
end if
%>
<% '-----------------修改一级导航--------------------------------------------------------------
if Request.QueryString("xiugai")="ok" then
id=request("id") 
sql="select * from menu where id="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("title")=request.form("title")
rs("en_title")=request.form("en_title") 
rs("url")=request.form("url") 
rs("openfs")=request.form("openfs") 
rs("px_id")=request.form("px_id") 
rs("yc")=request.form("yc") 
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""主菜单修改成功！"");window.location.href='admin_menu.asp?menu=admin';</script>")
end if
'-----------------修改二级导航--------------------------------------------------------------
if Request.QueryString("xiugai_nav")="ok" then
id=request("id") 
sql="select * from menu_nav where id="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("title")=request.form("title") 
rs("ssfl")=request.form("ssfl") 
rs("url")=request.form("url") 
rs("openfs")=request.form("openfs") 
rs("px_id")=request.form("px_id") 
rs("yc")=request.form("yc") 
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""二级菜单修改成功！"");window.location.href='admin_menu.asp?menu=admin';</script>")
end if

%> 
<%
if request("act")="del" then
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('admin_menu.asp?menu=admin');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from menu where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');document.location.href('admin_menu.asp?menu=admin');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	response.Write("<script language=""javascript"">alert(""主导航删除成功！"");window.location.href='admin_menu.asp?menu=admin';</script>")
end if
end if
'--------------二级分类删除-----------------------------------------------------------
if request("act")="nav_del" then
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('admin_menu.asp?menu=admin');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from menu_nav where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');window.location.href('admin_menu.asp?menu=admin');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	response.Write("<script language=""javascript"">alert(""二级导航删除成功！"");window.location.href='admin_menu.asp?menu=admin';</script>")
end if
end if
%>