<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="page.asp" -->
<%call chkAdmin(24)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>简历管理</title>
<script language="javascript"> 
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
<%if request.querystring("action")="admin" then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">简历管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
    
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="stable">
<form id="form1" name="form1" method="post" action="?"> 
<thead><tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox"/></td>
<td width="4%" align="center" class="td">ID</td>
<td width="17%" height="25" align="center" class="td">姓名</td>
<td width="16%" align="center" class="td">职位</td>
<td width="17%" align="center" class="td">姓 别</td>
<td width="11%" align="center" class="td">学历</td>
<td width="18%" align="center" class="td">时间</td>
<td width="7%" align="center" class="td">查看</td>
<td width="7%" align="center" class="td">删除</td>
</tr></thead>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from Resume order by id desc" 
rs.open exec,conn,1,1 
if rs.eof then
response.write ("<div style=""padding:10px;"">还没有收到任何简历!</div>")
else
rs.PageSize =30 '每页记录条数
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
<tr>
<td width="3%" align="center" class="td"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
<td width="4%" align="center" class="td"><input type="text" readonly style="text-align:center;width:40px" value="<%=rs("id")%>"/></td>
<td width="17%" height="25" align="center" class="td"><a href="admin_Resume.asp?action=ck&id=<%=rs("id")%>" style="color:#003399"><%=rs("name")%></a></td>
<td width="16%" align="center" class="td"><%=rs("ypzw")%></td>
<td width="17%" align="center" class="td"><%=rs("sex")%></td>
<td width="11%" align="center" class="td"><%=rs("xueli")%></td>
<td width="18%" align="center" class="td"><%=rs("data")%></td>
<td width="7%" align="center" class="td">
<input type="button" name="Submit3" value="查看" onclick="window.location.href='admin_Resume.asp?action=ck&id=<%=rs("id")%>' "  class="btn"/></td>
<td width="7%" align="center" class="td">

<input type="button" name="Submit" value="删除" onclick="javascript:if(confirm('确定删除？删除后不可恢复!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"  class="btn"/></td>
</tr>
<%rs.movenext 
if rs.eof then exit for 
next 
end if%>
<tr>
<td><input type="checkbox" name="allbox" onclick="CheckAll()" /></td>
<td align="center">全选</td>
<td colspan="7"><input type="submit" class="btn" onclick="form.action='?action=del';" value="删除"/><%call PageControl(iCount,maxpage,page,"border=0 align=right","<p align=right>")
rs.close
set rs=nothing
%></td>
</tr>
</form>
</table>        
   </td>
  </tr>
</table>
<%end if%>
<%if request.querystring("action")="ck" then
id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from Resume where id="&id
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">查看简历</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable" >
      <tr>
        <td width="11%" height="28" align="right" class="td">应聘职位：</td>
        <td width="89%"  class="td"><%=rs("ypzw")%></td>
      </tr>
      <tr>
        <td width="11%" height="25" align="right" class="td">姓名：</td>
        <td class="td"><%=rs("name")%></td>
      </tr>
      <tr>
        <td width="11%" height="25" align="right" class="td">性别：</td>
        <td class="td"><%=rs("sex")%></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">年龄：</td>
        <td class="td"><%=rs("nn")%></td>
    </tr>
      <tr>
        <td height="25" align="right" class="td">民族：</td>
        <td class="td"><%=rs("mz")%></td>
      </tr>
    <tr>
        <td height="25" align="right" class="td">户籍：</td>
        <td class="td"><%=rs("hj")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">婚姻状况：</td>
      <td class="td"><%=rs("hyzk")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">身高：</td>
      <td class="td"><%=rs("sg")%></td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">体重：</td>
      <td class="td"><%=rs("tz")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">身份证：</td>
      <td class="td"><%=rs("sfz")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">学历：</td>
      <td class="td"><%=rs("xueli")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">现所在地：</td>
      <td class="td"><%=rs("szd")%></td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">毕业院校：</td>
      <td class="td"><%=rs("byyx")%></td>
    </tr>
   <tr>
     <td height="25" align="right" class="td">联系电话：</td>
     <td class="td"><%=rs("tel")%></td>
   </tr>
   <tr>
     <td height="25" align="right" class="td">联系手机：</td>
     <td class="td"><%=rs("sj")%></td>
   </tr>
    <tr>
      <td height="25" align="right" class="td">教育背景：</td>
      <td class="td"><%=rs("jybj")%></td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">工作经验：</td>
      <td class="td"><%=rs("gzjn")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">专长：</td>
      <td class="td"><%=rs("zc")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">提交IP：</td>
      <td class="td"><%=rs("ip")%></td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">提交日期：</td>
      <td class="td"><%=rs("data")%></td>
    </tr>
   <tr>
      <td height="25" class="td"></td>
      <td class="td"><input type="button" name="Submit3" value="返回列表" onclick="window.location.href='admin_Resume.asp?action=admin' "  class="btn"/></td>
    </tr>
    
    </table></td>
  </tr>
</table>
<%end if%>
</body>
</html>
<%
if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from Resume where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('删除成功！');window.location.href='admin_Resume.asp';</script>"
end if

if Request.QueryString("action")="del" then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');history.go(-1);</script>" 
response.end()
end if
sql="delete from [case] where id in ("&Request("id")&")"
conn.execute(sql) 
Response.Write "<script>alert('恭喜!删除成功!');window.location.href='admin_Case.asp';</script>" 
end if

%>
