<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!--#include file="../Include/seeion.asp" -->
<% 
id=request.QueryString("pay_id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
set rsu=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rsu.open exec,conn,1,1 
userid=rsu("id")
username=rsu("zsname")
mail=rsu("mail")
qq=rsu("qq")
tel=rsu("tel")
set rs=server.createobject("adodb.recordset") 
exec="select * from [Products] where id="&id
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
<title>����:<%=rs("title")%>_<%=zych_home%></title>
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
<div class="bar"><b>���߶���<i>Order</i></b><span>��ǰλ�ã�<a href="/">��ҳ ></a><a href="/">���߶���</a></span></div>
<div class="w_650 fl">
<form name="alipayment" action="Payok.asp" method="get">
<table width="100%" border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="background:#FFF;border-collapse:collapse;line-height:35px; text-indent:3px">
<input name="cpid" value="<%=rs("id")%>" type="hidden" size=26  readonly>
<input name="userid" value="<%=userid%>" type="hidden" size=26  readonly>
<input name="WIDtitle" value="<%=rs("title")%>" type="hidden" size="30" readonly />
<input name="out_trade_no" value="<%=BuyId%>" type="hidden"  size="30" readonly />
<tr valign="middle">
  <td width="16%" align="right"class=field>��Ʒ����</td>
  <td width="84%"><a href="<%=dir%>Products/ShowProducts.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("img")%>" width="140" height="100" style="padding:4px;" /></a>
  <p><%=rs("title")%></p></td>
</tr>
<tr valign="middle">
  <td align="right" class=field>��Ʒ������</td>
  <td><%=BuyId%></td>
</tr>

<tr valign="middle">
  <td align="right" class=field>��Ʒ����</td>
  <td><input name="quantity" value="1" size="10"/> ��</td>
</tr>
<tr valign="middle">
  <td align="right" class=field>��Ʒ���</td>
  <td style="font-size:18px; color:#F00"><input name="WIDprice" type="hidden" value="<%=rs("zkj")%>" size="30" />��<span class="price"><%=rs("zkj")%>.00</span></td>
</tr>
<tr valign="middle">
  <td align="right" class=field>��Ʒչʾ��ַ</td>
  <td><input name="WIDshow_url" value="<%=zych_url%><%=dir%>Products/ShowProducts.asp?id=<%=rs("id")%>" size="30" readonly/></td>
</tr>
<tr valign="middle">
  <td align="right" class=field>�ջ�������</td>
  <td>
   <input name="WID_name" value="<%=username%>" size="30" />
   <span class=red>*</span><span class=kome>����</span> </td>
</tr>
<tr valign="middle">
  <td align="right" class=field>��ϵQQ</td>
  <td><input name="qq" value="<%=qq%>" size=30 maxlength=28 >
   <span class=red>*</span><span class=kome>����</span> </td>
</tr>
<tr valign="middle">
  <td align="right" class=field>E-mail����</td>
  <td><input name="email" value="<%=mail%>" size=30 maxlength=28 >
   <span class=red>*</span><span class=kome>���ڽ��ո�����Ϣ ��ʽ:aaaa@qq.com</span> </td>
</tr>
<tr valign="middle">
  <td align="right" class=field>�ջ����ֻ�����</td>
  <td>
  <input name="WID_tel" value="<%=tel%>" size="30" />
   <span class=red>*</span><span class=kome>����ȷ�ϸ�����Ϣ</span> </td>
</tr>
<tr valign="middle">
  <td align="right" class=field>�Ƿ񿪷�Ʊ</td>
  <td>
<script>
function setvalue(myitem){
if(myitem.checked){
myitem.value=1;
document.getElementById('obj').style.display="block"
document.getElementById("WIDtax").value =<%=rs("zkj")*rs("piaodian")%>;
document.getElementById("WIDFreight").value =<%=rs("kdf")%>;
document.getElementById("WIDtax_t").innerHTML =<%=rs("zkj")*rs("piaodian")%>;
document.getElementById("WIDFreight_t").innerHTML =<%=rs("kdf")%>;
}
else{
myitem.value=2;
document.getElementById('obj').style.display="none"
document.getElementById("WIDtax").value =0.00;
document.getElementById("WIDFreight").value =0.00; 
document.getElementById("WIDtax_t").innerHTML ="0.00";
document.getElementById("WIDFreight_t").innerHTML ="0.00";
}
}
</script>
<table width="300" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><input type="checkbox" id="test9" onClick="setvalue(this)" value="1" /></td>
    <td width="278"><span id="obj" style="display:none;">���ӷ��ã�˰��:<span id="WIDtax_t">0.00</span>Ԫ ���:<span id="WIDFreight_t">0.00</span>Ԫ</span>
   <input name="WIDtax" id="WIDtax" type="hidden"value="<%=rs("kdf")%>" />
   <input name="WIDFreight" id="WIDFreight" type="hidden" value="<%=rs("piaodian")%>"/></td>
  </tr>
</table>

</td>
</tr>
<tr>
  <td align="right" class=field>�ջ���ַ</td>
  <td><input name="WIDreceive_address" value="<%=rsu("gsadd")%>" size="30" />
   <span class=red>*</span><span class=kome>�磺XʡX��X��X·XС��X��X��ԪX��</span> </td>
</tr>
<tr valign="middle">
  <td align="right" class=field>&nbsp;</td>
  <td><span class="new-btn-login-sp"><button class="new-btn-login" type="submit" style="text-align:center;">ȷ ��</button></span>
    <font class="note-help">����������ȷ�ϡ���ť������ʾ��ͬ��ôε�ִ�в����� </font></td>
</tr>

</table>
</form>
</div>
   
<div class="w_280 fr">
<div class="box">
<div class="title tubiao_N">��Ա���� <span>User Nav</span></div>
<ul class="list_2">
<%=zych_user_nav%>
</ul>
</div>
<!--#include file="../Include/left.asp" -->
</div>
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
