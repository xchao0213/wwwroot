<!--#include file="../../Include/conn.asp" -->
<!--#include file="class/alipay_submit.asp" -->
<%id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [orders] where id="&id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('参数不正确，ID值不存在！');history.go(-1);</script>" 
Response.End()
end if%>
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title>支付宝付款_<%=zych_home%></title>
<link href="../images/style.css" rel="stylesheet" type="text/css" />
</head>
<body>

<%
'/////////////////////请求参数/////////////////////

'支付类型
payment_type = "1"
'必填，不能修改
'服务器异步通知页面路径
notify_url = ""&zych_url&"/alipay/danbao/notify_url.asp"
'需http://格式的完整路径，不能加?id=123这类自定义参数

'页面跳转同步通知页面路径
return_url = ""&zych_url&"/alipay/danbao/return_url.asp"
'需http://格式的完整路径，不能加?id=123这类自定义参数，不能写成http://localhost/

'卖家支付宝帐户
seller_email=alipayID
'必填
out_trade_no =rs("OrderNo") '商户订单号,订单系统中唯一订单号，必填
cpid=rs("cpid")'商品ID
userid=rs("userid")'网站用户ID
subject =rs("title")'订单名称
price = rs("rmb")'付款金额
quantity =rs("number")'商品数量'必填，建议默认为1，不改变值，把一次交易看成是一次下订单而非购买一件商品
logistics_fee =rs("kuaidi")'物流费用
logistics_type = "EXPRESS"'物流类型
'必填，三个值可选：EXPRESS（快递）、POST（平邮）、EMS（EMS）
'物流支付方式
logistics_payment = "BUYER_PAY"
'必填，两个值可选：SELLER_PAY（卖家承担运费）、BUYER_PAY（买家承担运费）

show_url =""&zych_url&"Products/ShowProducts.asp?id="&PID&""'商品展示地址
'需以http://开头的完整路径，如：http://www.xxx.com/myorder.html

receive_name =rs("name")'收货人姓名
qq=rs("qq")
quantity=rs("number")
email=rs("email")

receive_address =rs("address")'收货人地址
'如：XX省XXX市XXX区XXX路XXX小区XXX栋XXX单元XXX号

'收货人电话号码
receive_phone =rs("tel")
'如：0571-88158090

'收货人手机号码
receive_mobile =rs("tel")
'如：13312341234

'/////////////////////请求参数/////////////////////

'构造请求参数数组
sParaTemp = Array("service=create_partner_trade_by_buyer","partner="&partner,"seller_email="&seller_email,"_input_charset="&input_charset  ,"payment_type="&payment_type   ,"notify_url="&notify_url   ,"return_url="&return_url   ,"out_trade_no="&out_trade_no   ,"subject="&subject   ,"price="&price   ,"quantity="&quantity   ,"logistics_fee="&logistics_fee   ,"logistics_type="&logistics_type   ,"logistics_payment="&logistics_payment   ,"body="&body   ,"show_url="&show_url   ,"receive_name="&receive_name   ,"receive_address="&receive_address   ,"receive_zip="&receive_zip   ,"receive_phone="&receive_phone   ,"receive_mobile="&receive_mobile  )

'建立请求
Set objSubmit = New AlipaySubmit
sHtml = objSubmit.BuildRequestForm(sParaTemp, "get", "确认")
response.Write sHtml


%>
</body>
</html>
