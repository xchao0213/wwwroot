<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(15)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>产品一级分类管理</title>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">产品一级分类</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<thead><tr>
<td width="8%" align="center" class="td">ID</td>
<td width="27%" height="25" align="center" class="td">分类名称</td>
<td width="13%" align="center" class="td">所属分类</td>
<td width="28%" align="center" class="td">分类地址</td>
<td width="10%" align="center" class="td">排 序</td>
<td width="7%" align="center" class="td">修改</td>
<td width="7%" align="center" class="td">删除</td>
</tr></thead>
<%	
set rs=server.createobject("adodb.recordset") 
exec="select * from Bigclass order by px_id asc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">暂无一级分类!</div>")
else
rs.PageSize =300 '每页记录条数
iCount=rs.RecordCount '记录总数
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
<form action="?xiugai=BigClass" method="post" name="add">
<tr style="background: #F5F5F5">
<td width="8%" align="center" class="td"><%=rs("BigClassID")%><input name="BigClassID" type="hidden" size="15"  value="<%=rs("BigClassID")%>"/></td>
<td width="27%" height="25" class="td"><input name="BigClassName" type="text" size="15"  value="<%=rs("BigClassName")%>"/></td>
<td width="13%" align="center" class="td">一级分类</td>
<td width="28%" class="td">/Products/?BigClassID=<%=rs("BigClassID")%></td>
<td width="10%" class="td"><input name="px_id" type="text" value="<%=rs("px_id")%>" size="7"  /></td>
<td width="7%" align="center" class="td"><input type="submit" name="button2" id="button2" value="修改"  class="btn"/></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_products_bigclass.asp?act=del_BigClass&id=<%=rs("BigClassID")%>';}else{history.go(0);}"  class="btn"/></td>
</tr></form>
<%set rs2=server.createobject("adodb.recordset") 
exec="select * from [SmallClass] where BigClassID="&rs("BigClassID")&" order by px_id asc  " 
rs2.open exec,conn,1,1
if rs2.eof and rs2.bof then
response.Write("")
else
do while not rs2.eof%>
<form action="?xiugai=SmallClass" method="post" name="add">
<tr>
<td width="8%" align="center" class="td"><input name="SmallClassID" type="hidden" size="15"  value="<%=rs2("SmallClassID")%>"/></td>
<td width="27%" class="td" style="padding-left:20px">├─<%=rs2("px_id")%>.<input name="SmallClassName" type="text" size="15"  value="<%=rs2("SmallClassName")%>"/></td>
<td width="13%" align="center" class="td">├─<select name="BigClassID" id="select">
<%
dim rsc
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from bigclass",conn,1,1
while not rsc.eof
if rs2("BigClassID")=rsc("BigClassID") then
response.Write("<option value=""" & rsc("BigClassID") & """ selected>" & rsc("BigClassName") & "</option>")
else
response.Write("<option value=""" & rsc("BigClassID") & """>" & rsc("BigClassName") & "</option>")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
%>
</select>
</td>
<td width="28%" class="td">├─/Products/?BigClassID=<%=rs("BigClassID")%>&SmallClassID=<%=rs2("SmallClassID")%></td>
<td width="10%" class="td">├─<input name="px_id" type="text" value="<%=rs2("px_id")%>" size="4"  /></td>
<td width="7%" align="center" class="td"><input type="submit" name="button2" id="button2" value="修改"  class="btn"/></td>
<td width="7%" align="center" class="td"><input type="button" name="Submit" value="删除" onClick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='admin_products_BigClass.asp?act=del_SmallClass&id=<%=rs2("SmallClassID")%>';}else{history.go(0);}"  class="btn"/></td>
</tr></form>
<%
rs2.movenext
loop
end if
rs2.close
set rs2=nothing

rs.movenext 
if rs.eof then exit for 
next 
end if
%>

</table>
    </td>
  </tr>
</table>

<div style="margin-top:10px">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?add=BigClass" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加一级分类</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="18%" height="25" align="right" class="td">一级分类名称</td>
        <td width="82%"  class="td"><input name="BigClassName" type="text" size="30"  /></td>
      </tr>
      
      <tr>
        <td width="18%" height="25" align="right" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /> 
          数字越小越靠前。</td>
      </tr>
      
      <tr>
        <td height="25" align="right" class="td">&nbsp;</td>
        <td class="td"><label>
        <input type="submit" name="button" id="button" value="增加分类"  class="btn"/>
        </label></td>
      </tr>
      
    </table></td>
  </tr></form>
</table>

    </div>
<div style="margin-top:10px">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?add=SmallClass" method="post" name="add">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">增加二级分类</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr >
        <td width="18%" height="13" align="right" class="td">二级分类名称</td>
        <td width="82%"  class="td"><input name="SmallClassName" type="text" size="30"  /></td>
      </tr>
      <tr >
        <td width="18%" height="12" align="right" class="td">所属大类</td>
        <td  class="td"><select name="BigClassID" id="select">
<%
dim rss
set rss=server.CreateObject("adodb.recordset")
rss.open "select * from bigclass",conn,1,1
while not rss.eof
response.Write("<option value=""" & rss("BigClassID") & """>" & rss("BigClassName") & "</option>")
rss.movenext
wend
rss.close
set rss=nothing
%></select></td>
      </tr>
      
      <tr>
        <td width="18%" height="25" align="right" class="td">排序ID</td>
        <td class="td"><input name="px_id" type="text" size="30"  /> 
          数字越小越靠前。</td>
      </tr>
      
      <tr>
        <td height="25" class="td">&nbsp;</td>
        <td class="td"><label>
        <input type="submit" name="button" id="button" value="增加分类"  class="btn"/>
        </label></td>
      </tr>
      
    </table></td>
  </tr></form>
</table>
</div>
</body>
</html>
<% 
if Request.QueryString("xiugai")="BigClass" then
id=request("BigClassID") 
sql="select * from bigclass where BigClassID="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("BigClassName")=request.form("BigClassName") 
rs("px_id")=request.form("px_id") 
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""当前分类修改成功！"");window.location.href='admin_products_bigclass.asp';</script>")
end if
'删除一级分类
if request("act")="del_BigClass" then
	id=request("ID")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');history.go(-1);</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from bigclass where BigClassID="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');history.go(-1);</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('当前分类删除成功！');window.location.href='admin_products_BigClass.asp';</script>"
end if
end if
'新增一级分类
if request("add")="BigClass" then
set rs=server.createobject("adodb.recordset")
sql="select * from bigclass"
rs.open sql,conn,1,3
BigClassName=request.form("BigClassName")
px_id=request.form("px_id")
if BigClassName=""  then 
response.Write("<script language=javascript>alert('分类名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert(""错误！排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("BigClassName")=BigClassName
rs("px_id")=px_id
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('一级分类增加成功！');window.location.href='admin_products_bigclass.asp';</script>" 
end if
%>

<% '二级分类修改
if Request.QueryString("xiugai")="SmallClass" then
id=request("SmallClassID") 
sql="select * from smallclass where SmallClassID="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
IF not isNumeric(request("px_id")) then
response.write("<script>alert(""排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("SmallClassName")=request.form("SmallClassName") 
rs("BigClassID")=request.form("BigClassID") 
rs("px_id")=request.form("px_id") 
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""当前分类修改成功！"");window.location.href='admin_products_BigClass.asp';</script>")
end if
'删除二级分类
if request("act")="del_SmallClass" then
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');history.go(-1);</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from  smallclass where SmallClassID="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');history.go(-1);</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('当前分类删除成功！');window.location.href='admin_products_BigClass.asp';</script>"
end if
end if
'新增二级分类
if request("add")="SmallClass" then
set rs=server.createobject("adodb.recordset")
sql="select * from smallclass"
rs.open sql,conn,1,3
SmallClassName=request.form("SmallClassName")
BigClassID=request.form("BigClassID")
px_id=request.form("px_id")
if SmallClassName=""  then 
response.Write("<script language=javascript>alert('二级分类名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if px_id=""  then 
response.Write("<script language=javascript>alert('排序ID不能为空!');history.go(-1)</script>") 
response.end 
end if
IF not isNumeric(request("px_id"))  then
response.write("<script>alert("" 错误！排序ID必须为数字！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("SmallClassName")=SmallClassName
rs("BigClassID")=BigClassID
rs("px_id")=px_id
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
Response.Write "<script>alert('二级分类增加成功！');window.location.href='admin_products_BigClass.asp';</script>" 
end if
%>
