<!--#include file="../../Include/conn.asp" -->
<%
' 功能：支付宝页面跳转同步通知页面' 版本：3.3' 日期：2012-07-17%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="class/alipay_notify.asp"-->
<title>支付宝支付结果</title>
<meta http-equiv="refresh" content="3;url=<%=zych_url&"/User/?action=admin"%>">
</head>
<style>
.mbox{width:600px; height:200px; position:absolute; left:50%; top:50%; margin-left:-300px; margin-top:-100px;background:#edffcd; border:1px #9bdb51 solid}
.mbox .cgbox{ height:80px;border-bottom:1px #CCCCCC solid}
.mbox .cgbox p{width:380px; padding-left:50px; font-size:30px;line-height:42px; margin:40px auto 20px auto; background:url(../images/gxico.png) no-repeat;}
.mbox .xwbox{ height:80px;border-bottom:1px #CCCCCC solid}
.mbox .xwbox p{width:380px; padding-left:50px; font-size:30px;line-height:42px; margin:40px auto 20px auto; background:url(../images/gxico.png) 0px -42px no-repeat;}
.mbox .body{ padding:10px; text-align:center; font-size:12px;}
.mbox .body span{font-weight:bold; padding-right:5px}
</style>
<body>
<%
'计算得出通知验证结果
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyReturn()

If sVerifyResult Then	'验证成功
	'*********************************************************************
	'请在这里加上商户的业务逻辑程序代码
	
	'——请根据您的业务逻辑来编写程序（以下代码仅作参考）——
    '获取支付宝的通知返回参数，可参考技术文档中页面跳转同步通知参数列表

	'商户订单号

	out_trade_no = Request.QueryString("out_trade_no")

	'支付宝交易号

	trade_no = Request.QueryString("trade_no")

	'交易状态
	trade_status = Request.QueryString("trade_status")

	
	If Request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" Then
	'判断是否在商户网站中已经做过了这次通知返回的处理
		'如果没有做过处理，那么执行商户的业务程序
		'如果有做过处理，那么不执行商户的业务程序
		conn.execute("update orders set state='2' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		
	ElseIf request.QueryString("trade_status") = "TRADE_FINISHED" Then '完成交易
		'判断该笔订单是否在商户网站中已经做过处理
		'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
		'如果有做过处理，不执行商户的业务程序
		conn.execute("update orders set state='8' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
	Else
		Response.Write "trade_status="&Request.QueryString("trade_status")
	End If

	'Response.Write "验证成功<br>"
		Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""cgbox""><p>您的订单，支付成功了！</p></div>"
		Response.Write"<div class=""body""><span>提示：</span>您已经成功付款,正在处理您的订单请勿关闭此自页面。处理完毕自动跳转！</div>"
		Response.Write"</div>"
	'——请根据您的业务逻辑来编写程序（以上代码仅作参考）——
	
	'*********************************************************************
else '验证失败
    '如要调试，请看alipay_notify.asp页面的VerifyReturn函数，比对sign和mysign的值是否相等，或者检查responseTxt有没有返回true
    response.Write "验证失败"
end if
%>
</body>
</html>
