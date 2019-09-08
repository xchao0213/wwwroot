<!--#include file="../../Include/conn.asp" -->
<!--#include file="class/alipay_submit.asp"-->
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title>支付宝付款_<%=zych_home%></title>
<link href="../images/main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<%
'支付类型
payment_type = "1"  '必填，不能修改

'服务器异步通知页面路径
notify_url = ""&zych_url&"/alipay/Guarantee/notify_url.asp"
'需http://格式的完整路径，不能加?id=123这类自定义参数

'页面跳转同步通知页面路径
return_url = ""&zych_url&"/alipay/Guarantee/return_url.asp"
'需http://格式的完整路径，不能加?id=123这类自定义参数，不能写成http://localhost/

seller_email=alipayID'卖家支付宝帐户

out_trade_no = Request.Form("out_trade_no")'商户订单号,订单系统中唯一订单号，必填
'商户网站
cpid=Request.Form("cpid")

userid=Request.Form("userid")
'订单名称
subject = Request.Form("WIDtitle")

'付款金额
price = cint(Request.Form("WIDprice"))+cint(Request.Form("WIDtax"))
'price = Request.Form("WIDprice")

quantity = Request.Form("quantity") '商品数量
'必填，建议默认为1，不改变值，把一次交易看成是一次下订单而非购买一件商品
'物流费用
logistics_fee =Request.Form("WIDFreight")
'必填，即运费
'物流类型
logistics_type = "EXPRESS"
'必填，三个值可选：EXPRESS（快递）、POST（平邮）、EMS（EMS）
'物流支付方式
logistics_payment = "BUYER_PAY"
'必填，两个值可选：SELLER_PAY（卖家承担运费）、BUYER_PAY（买家承担运费）

'商品展示地址
show_url = Request.Form("WIDshow_url")
'需以http://开头的完整路径，如：http://www.xxx.com/myorder.html

'收货人姓名
receive_name = Request.Form("WID_name")
qq= Request.Form("qq")
web= Request.Form("WIDweb")
email= Request.Form("email")
'如：张三

'收货人地址
receive_address = Request.Form("WIDreceive_address")
'如：XX省XXX市XXX区XXX路XXX小区XXX栋XXX单元XXX号

'收货人电话号码
receive_phone = Request.Form("WIDreceive_phone")
'如：0571-88158090

'收货人手机号码
receive_mobile = Request.Form("WID_tel")
'如：13312341234

'/////////////////////请求参数/////////////////////

'构造请求参数数组
sParaTemp = Array("service=trade_create_by_buyer","partner="&partner,"_input_charset="&input_charset  ,"payment_type="&payment_type   ,"notify_url="&notify_url   ,"return_url="&return_url   ,"seller_email="&seller_email   ,"out_trade_no="&out_trade_no   ,"subject="&subject   ,"price="&price   ,"quantity="&quantity   ,"logistics_fee="&logistics_fee   ,"logistics_type="&logistics_type   ,"logistics_payment="&logistics_payment   ,"body="&body   ,"show_url="&show_url   ,"receive_name="&receive_name   ,"receive_address="&receive_address   ,"receive_phone="&receive_phone   ,"receive_mobile="&receive_mobile  )

'建立请求
Set objSubmit = New AlipaySubmit
sHtml = objSubmit.BuildRequestForm(sParaTemp, "get", "确认")
response.Write sHtml
%>
</body>
</html>
