<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!--#include file="../Include/seeion.asp" -->
<!--#include file="../Include/page.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ա����_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>��Ա���� <i>User</i></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>User/">��Ա����</a>�ҵĶ���</span></div>
<div class="w_650 fl">
<div class="position">
<table width="100%" border="1" cellpadding="10" cellspacing="0" bordercolor="#EDEDED" Class="fkzt yh">
  <tr>
    <td width="25" align="center">1.�ȴ���Ҹ���</td>
    <td width="25" align="center">2.�Ѹ���,�ȴ�����</td>
    <td width="25" align="center">3.�ѷ���,�ȴ�ȷ���ջ�</td>
    <td width="25" align="center">4.ȷ���ջ�,�������</td>
  </tr>
</table>
<div class="ubox">
<div class="pronav"><h1>�������</h1></div>


<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse: collapse;line-height:25px; text-indent:3px">
<tr>
    <td width="19%" align="center"><strong>�������</strong></td>
    <td width="27%" align="center"><strong>��Ʒ����</strong></td>
    <td width="23%" align="center"><strong>ʱ ��</strong></td>
    <td width="19%" align="center"><strong>״ ̬</strong></td>
    <td colspan="2" align="center"><strong>�� ��</strong></td>
    </tr>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from [orders] where userid="&session("username")&" order by id desc "  
rs.open exec,conn,1,1 
if rs.eof then
response.Write "&nbsp;���޶�����"
else
rs.PageSize =15 'ÿҳ��¼����
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
For i=1 To x
if rs("state")=1 then
ddzt="<a href="""&dir&"alipay/"&alipayLX&"/alipayapi_re.asp?id="&rs("id")&""" title=""���������ɸ���"" class=""kfal_a"" target=""_blank"" class=""kfal_a"" onClick=""document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'""><font>��������</font></a>"
elseif rs("state")=2 then
ddzt="����Ѹ���ȴ�����"
elseif rs("state")=3 then
ddzt="�ѷ������ȴ�ȷ���ջ�"
elseif rs("state")=8 then
ddzt="�������"
else
ddzt="���������쳣"
end if%>
  <tr>
    <td width="19%" align="center"><a href="my_Showorders.asp?id=<%=rs("id")%>" title="����鿴��������"><%=rs("OrderNo")%></a></td>
    <td width="27%"><%=stvalue(rs("title"),24)%></td>
    <td width="23%" align="center"><%=rs("data")%></td>
    <td width="19%" align="center"><%=ddzt%></td>
    <td width="6%" align="center"><a href="my_Showorders.asp?id=<%=rs("id")%>">�鿴</a></td>
    <td width="6%" align="center"><a href="javascript:" onClick="javascript:if(confirm('ȷ��ɾ���ö�����ɾ���󲻿ɻָ�!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}" title="ɾ���˶���">ɾ��</a></td>
  </tr>
<%rs.movenext
next
end if%>
</table>
<%'������ʾ��ҳ
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%>
</div></div>
</div>
   
<div class="w_280 fr">
    <div class="box">
      <div class="title tubiao_U">��Ա���� <span>User Nav</span></div>
      <ul class="list_2">
        <%=zych_user_nav%>
      </ul>
    </div>
    <!--#include file="../Include/left.asp" -->
  </div>
</div>
<div id="light" class="Order_tc">
<span class="ortitle">����֧����ʾ</span>
<div class="up-box">
<span class="icon"><img src="../images/working.gif"></span>
<div class="r-txt">֧�����ǰ���벻Ҫ�رմ�֧����֤���ڡ�<br clear="none">֧����ɺ��������֧�������������水ť��</div>
</div> 
<div class="lay-btn">
<a href="/User" class="btn92" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">֧����������</a>
<a href="/User/my_orders.asp" class="btn92s" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'" shape="rect">֧�����</a>
</div>
</div> 
<div id="fade" class="Order_bj"> </div>
<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<%if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from orders where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰ����ɾ���ɹ���');window.location.href='my_orders.asp';</script>"
end if %>