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
WID_tel=request.QueryString("phone")
WID_rmb=request.QueryString("rmb")
WIDreceive_address=request.QueryString("WIDreceive_address")
qq=request.QueryString("qq")
WIDtax=request.QueryString("WIDtax")
WIDFreight=request.QueryString("WIDFreight")
if qq=""  then 
response.Write("<script language=javascript>alert('联系QQ不能为空');history.go(-1)</script>") 
response.end 
end if
if WID_tel=""  then 
response.Write("<script language=javascript>alert('收货人手机号码不能为空');history.go(-1)</script>") 
response.end 
end if
if WIDreceive_address=""  then 
response.Write("<script language=javascript>alert('收货地址不能为空');history.go(-1)</script>") 
response.end 
end if
set rsu=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rsu.open exec,conn,1,1 
userid=rsu("id")
username=rsu("zsname")
mail=rsu("mail")
randNumber=getStrRandNumber(1000,9999)
BuyId=""&getTime&randNumber&""%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>确认购买_<%=WIDtitle%></title>
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
<div class="bar"><b>确认订单信息</b><span>当前位置：<a href="/">首页 ></a><a href="/">确认订单</a>> <a href="/"><%=WIDtitle%></a></span></div>
<div class="ordok">
<form name=alipayment action="<%=alipayLX%>/alipayapi.asp" method="post" target="_blank">
<table width="100%"  border="1" cellpadding="10" cellspacing="0" bordercolor="#DeDeDe">
<tr>
<td colspan="2" align="center">产品名称</td>
<td width="20%" align="center">单价（元）</td>
<td width="16%" align="center">数量</td>
<td width="24%" align="center">金额</td>
</tr>
<% set rs=server.createobject("adodb.recordset") 
exec="select * from [Products] where id in ("&cpid&")"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此商品</div>"
response.End()
end if
while not rs.eof%>
<tr>
<td width="13%"><img src="<%=rs("img")%>" width="100" height="60" style="padding:4px;" /></td>
<td width="27%"><a href="<%=dir%>Products/ShowProducts.asp?id=<%=rs("id")%>" target="_blank"><%=rs("title")%></a><input name="WIDtitle" value="<%=rs("title")%>" type="hidden" size="30"/></td>
<td width="20%" align="center"><%=rs("zkj")%></td>
<td width="16%" align="center"><% set rsp=server.createobject("adodb.recordset") 
exec="select * from [ogwc] where sfdd=0 and userid="&session("username")&" and cpid="&rs("id")&""
rsp.open exec,conn,1,1 
if rsp.eof and rsp.bof then
response.Write""
else
cp_cpsl=rsp("cpsl")
cp_rmb=rs("zkj")*rsp("cpsl")
end if
rsp.close
set rsp=nothing %>
<%=cp_cpsl%></td>
<td width="24%" align="center"style="color:#F50;font: 700 12px tahoma;"><%=cp_rmb%></td>
</tr>
<% rs.movenext
wend
rs.close
set rs=nothing %>
<tr valign="middle">
  <td colspan="2" rowspan="3" class=field>&nbsp;</td>
  <td colspan="3" class=field>附加费用：税金:<span id="WIDtax_t"><%= WIDtax %></span>元 快递:<span id="WIDFreight_t"><%=WIDFreight %></span>元
    <input name="WIDtax" id="WIDtax" type="hidden"value="<%=WIDtax%>" />
    <input name="WIDFreight" id="WIDFreight" type="hidden" value="<%=WIDFreight %>"/></td>
</tr>

<tr>
  <td class=field>实付金额：￥<span class="price"><%=WID_rmb%>.00</span></td>
  <td colspan="2" align="right">
    <p><b>收货地址:</b> <%=WIDreceive_address%></p>
    <p><b>收货人：</b><%=WID_name%> <%=WID_tel%> <b>QQ：</b><%=qq%></p></td>
  </tr>
<tr valign="middle">
  <td align="center"><a href="/User/my_addlist.asp">返回购物车</a></td>
<input name="WID_name" type="hidden"  value="<%=WID_name%>" size="30" />
<input name="WID_tel" type="hidden"  value="<%=WID_tel%>" size="30" />
<input name="qq" type="hidden"  value="<%=qq%>" size=30 maxlength=28 />
<input name="email" value="<%=email%>" type="hidden" size=30 maxlength=28 />
<input name="quantity" type="hidden"  value="<%= quantity %>" size="10"/>
<input name="WIDprice" type="hidden" value="<%=WID_rmb%>" size="30" />
<input name="WIDreceive_address" type="hidden"  value="<%=WIDreceive_address%>" size="30" />
<input name="WIDshow_url" type="hidden"  value="<%=WIDshow_url%>" size="30"/>
<input name="cpid" value="<%=cpid%>" type="hidden" size=26 >
<input name="userid" value="<%=userid%>" type="hidden" size=26 >
<input name="out_trade_no" value="<%=BuyId%>" type="hidden"  size="30"/>
  <td colspan="2" align="right"><span class="new-btn-login-sp"><button class="new-btn-login" type="submit" onClick="document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'">确认支付</button></span></td>

</tr>

</table>
</form>
</div>
</div>
<div id="light" class="Order_tc">
<span class="ortitle">网上支付提示</span>
<div class="up-box">
<span class="icon"><img src="../images/working.gif"></span>
<div class="r-txt">支付完成前，请不要关闭此支付验证窗口。<br clear="none">支付完成后，请根据您支付的情况点击下面按钮。</div>
</div> 
<div class="lay-btn">
<a href="/User" class="btn92" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">支付遇到问题</a>
<a href="/User/my_orders.asp" class="btn92s" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'" shape="rect">支付完成</a>
</div>
</div> 
<div id="fade" class="Order_bj"> </div>
<!--#include file="../Include/bottom.asp" -->
</BODY>
</HTML>