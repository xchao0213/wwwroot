<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/seeion.asp"--> 
<% 
cpid=request.QueryString("cpid")
userid=request.QueryString("userid")
WIDtitle=request.QueryString("WIDtitle")
out_trade_no=request.QueryString("out_trade_no")
quantity=request.QueryString("quantity")
WIDprice=request.QueryString("WIDprice")
WIDshow_url=request.QueryString("WIDshow_url")
WID_name=request.QueryString("WID_name")
email=request.QueryString("email")
WID_tel=request.QueryString("WID_tel")
WIDreceive_address=request.QueryString("WIDreceive_address")
qq=request.QueryString("qq")
WIDtax=request.QueryString("WIDtax")
WIDFreight=request.QueryString("WIDFreight")
if qq=""  then 
response.Write("<script language=javascript>alert('��ϵQQ����Ϊ��');history.go(-1)</script>") 
response.end 
end if
if email=""  then 
response.Write("<script language=javascript>alert('��ϵ���䲻��Ϊ��');history.go(-1)</script>") 
response.end 
end if
if WID_tel=""  then 
response.Write("<script language=javascript>alert('�ջ����ֻ����벻��Ϊ��');history.go(-1)</script>") 
response.end 
end if
if WIDreceive_address=""  then 
response.Write("<script language=javascript>alert('�ջ���ַ����Ϊ��');history.go(-1)</script>") 
response.end 
end if
set rsu=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rsu.open exec,conn,1,1 
userid=rsu("id")
username=rsu("zsname")
mail=rsu("mail")
set rs=server.createobject("adodb.recordset") 
exec="select * from [Products] where id="&cpid
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">�޴���Ʒ</div>"
response.End()
end if
randNumber=getStrRandNumber(1000,9999)
BuyId=""&getTime&randNumber&""%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ȷ�Ϲ���_<%=WIDtitle%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
</head>
</HEAD>
<BODY>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>ȷ�϶���</b><span>��ǰλ�ã�<a href="/">��ҳ ></a><a href="/">ȷ�϶���</a>> <a href="/"><%=WIDtitle%></a></span></div>
<div class="ordok">
<form name=alipayment action="<%=alipayLX%>/alipayapi.asp" method="post" target="_blank">
<table width="100%"  border="1" cellpadding="10" cellspacing="0" bordercolor="#D8D8D8">
<input name="cpid" value="<%=cpid%>" type="hidden" size=26 >
<input name="userid" value="<%=userid%>" type="hidden" size=26 >
<input name="WIDtitle" value="<%=WIDtitle%>" type="hidden" size="30"/>
<input name="out_trade_no" value="<%=BuyId%>" type="hidden"  size="30"/>
<tr valign="middle">
  <td width="10%" rowspan="4" class=field>��Ʒ����</td>
  <td width="22%" rowspan="4" align="center"><a href="<%=dir%>Products/ShowProducts.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("img")%>" width="160" height="100" style="padding:4px;" /></a>
  <p><%=WIDtitle%>
    <input name="WIDshow_url" type="hidden"  value="<%=WIDshow_url%>" size="30"/>
  </p></td>
  <td width="34%"><span class="field">���ţ�</span><%=BuyId%></td>
  <td width="34%"><span class="field">�ջ��ˣ�</span>
    <input name="WID_name" type="hidden"  value="<%=WID_name%>" size="30" /><%=WID_name%></td>
</tr>
<tr valign="middle">
  <td><span class="field">������</span>
    <input name="quantity" type="hidden"  value="<%= quantity %>" size="10"/> <%=quantity%>��</td>
  <td width="34%">�� �䣺
    <input name="email" value="<%=email%>" type="hidden" size=30 maxlength=28 /><%=email%></td>
</tr>
<tr valign="middle">
  <td>��
    <input name="WIDprice" type="hidden" value="<%=WIDprice%>" size="30" /><span class="price"><%=WIDprice%></span>��</td>
  <td width="34%">�� ����
    <input name="WID_tel" type="hidden"  value="<%=WID_tel%>" size="30" /><%=WID_tel%></td>
</tr>
<tr valign="middle">
  <td>���ӷ��ã�˰��:<span id="WIDtax_t"><%= WIDtax %></span>Ԫ ���:<span id="WIDFreight_t"><%=WIDFreight %></span>Ԫ
    <input name="WIDtax" id="WIDtax" type="hidden"value="<%= WIDtax %>" />
    <input name="WIDFreight" id="WIDFreight" type="hidden" value="<%=WIDFreight %>"/></td>
  <td width="34%">Q Q��
    <input name="qq" type="hidden"  value="<%=qq%>" size=30 maxlength=28 /><%=qq%></td>
</tr>
<tr>
  <td class=field>�ջ���ַ</td>
  <td colspan="3"><input name="WIDreceive_address" type="hidden"  value="<%=WIDreceive_address%>" size="30" /><%=WIDreceive_address%></td>
</tr>
<tr valign="middle">
  <td class=field>&nbsp;</td>
  <td colspan="3"><span class="new-btn-login-sp"><button class="new-btn-login" type="submit" onClick="document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'">ȷ��֧��</button></span>
    <font class="note-help">����������ȷ�ϡ���ť������ʾ��ͬ��ôε�ִ�в����� </font></td>
</tr>

</table>
</form>
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
</BODY>
</HTML>