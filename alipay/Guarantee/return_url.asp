<!--#include file="../../Include/conn.asp" -->
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<%
' 功能：支付宝页面跳转同步通知页面
' 版本：3.3
' 日期：2012-07-17
' 说明：
' 以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
' 该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
' //////////////页面功能说明//////////////
' 该页面可在本机电脑测试
' 可放入HTML等美化页面的代码、商户业务逻辑程序代码
' 该页面可以使用ASP开发工具调试，也可以使用写文本函数LogResult进行调试，该函数已被默认关闭，见alipay_notify.asp中的函数VerifyReturn
' （该接口的注意事项，如果没有那么该行删除）。
'////////////////////////////////////////
%>
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

If sVerifyResult Then'验证成功
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

	If Request.QueryString("trade_status") ="WAIT_SELLER_SEND_GOODS" Then '买家已付款，等待卖家发货
	    '判断是否在商户网站中已经做过了这次通知返回的处理
		'如果没有做过处理，那么执行商户的业务程序
		'如果有做过处理，那么不执行商户的业务程序
		conn.execute("update orders set state='2' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		SendEmail ""&email&"","您好，您正在买:"&title&"","您好，你已经购了本站商品："&title&"，我们将尽快给你发货，","Jmail" 

	ElseIf request.QueryString("trade_status") = "TRADE_FINISHED" Then '完成交易
		'判断该笔订单是否在商户网站中已经做过处理
		'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
		'如果有做过处理，不执行商户的业务程序
		conn.execute("update orders set state='8' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		
		set rs=server.createobject("adodb.recordset") 
		exec="select * from [orders] where OrderNo='"&out_trade_no&"'" 
		rs.open exec,conn,1,1 
		if rs.eof and rs.bof then
		Response.Write "d" 
		else
		title=rs("title")
		rmb=rs("rmb")
		email=rs("email")
		proid=rs("cpid")
		end if
		Function body()
		body=body+"<style type=""text/css"">"
		body=body+"<!--"
		body=body+".qmbox{margin:10px}"
		body=body+".qmbox p{line-height:15px;color:#666}"
		body=body+".qmbox p span{color:#F60}"
		body=body+"-->"
		body=body+"</style>"
		body=body+"<div class=qmbox>"
		body=body+"<div style=""width:700px;border-bottom:1px solid #ccc;margin-bottom:30px;"">"
		body=body+"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""700"" height=""39"" style=""font:12px Tahoma, Arial, 宋体;"">"
		body=body+"<tbody><tr>"
		body=body+"<td width=""210""><a target=_blank href="&zych_url&"><img src="""&zych_url&"/images/logo.png"""" height=""39"" border=""0""></a></td>"
		body=body+"<td width=""490"" align=""right"" valign=""bottom"" style=""padding-bottom:10px;""><a target=""_blank"" style=""color:#07f;text-decoration:none;"" href="""&zych_url&"/Member/"">会员中心</a></td>"
		body=body+"</tr></tbody></table></div>"
		body=body+"<div style=""width:680px;padding:0 10px;"">"
		body=body+"<div style=""line-height:1.5;font-size:14px;margin-bottom:25px;color:#4d4d4d;""><strong style=""display:block;margin-bottom:15px;"">亲爱的会员您好！</strong>"
		body=body+"<p>您的已于 "&Request.QueryString("notify_time")&" 成功在 "&zych_url&" 购买了 ：<strong style=""font-size:14px;""><a target=_blank style=color:#f60; href="&zych_url&dir&"Products/ShowProducts.asp?id="&proid&">"&title&"</a></strong></p></div>"
		body=body+"<div style=""margin-bottom:30px;"">"
		body=body+"<p>商品订单：<span>"&Request.QueryString("out_trade_no")&"</span></p>"
		body=body+"<p>付款金额：<span>"&Request.QueryString("total_fee")&"元</span></p>"
		body=body+"<p>邮寄费用：<span>"&Request.QueryString("logistics_fee")&"元</span></p>"
		body=body+"<p>交易状态：<span>"&Request.QueryString("trade_status")&" (交易完成)</span></p>"
		body=body+"<p>支付宝交易号：<span>"&Request.QueryString("trade_no")&"</span></p>"
		body=body+"<p>如果您在交易过程中遇到什么情况以上信息可作为付款凭证。</p>"
		body=body+"<a href="&zych_url&dir&"Products/ShowProducts.asp?id="&proid&" target=_blank>"&zych_url&dir&"Products/ShowProducts.asp?id="&proid&"</a></div>"
		body=body+"<div style=""margin-bottom:30px;"">"
		body=body+"<p>"&zych_description&"</p>"
		body=body+"<p><a href="&zych_url&" target=_blank>"&zych_url&"</a></p>"
		body=body+"</div></div>"
		body=body+"<div style=""padding:10px 10px 0;border-top:1px solid #ccc;color:#999;margin-bottom:20px;line-height:1.3em;font-size:12px;"">"
		body=body+"<p style=""margin-bottom:15px;"">此为系统邮件，请勿回复<br>请保管好您的邮箱，避免支付宝账户被他人盗用<br>Copyright zych 2010-2014 All Right Reserved</p>"
		body=body+"</div>  </div>"
		End Function
		SendEmail ""&email&"","您已经购买了:"&title&"",""&body&"","Jmail" 
	Else
		Response.Write "trade_status="&Request.QueryString("trade_status")
	End If
		Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""cgbox""><p>您的订单，支付成功了！</p></div>"
		Response.Write"<div class=""body""><span>提示：</span>您已经成功付款,正在处理您的订单请勿关闭此自页面。处理完毕自动跳转！</div>"
		Response.Write"</div>"
	'——请根据您的业务逻辑来编写程序（以上代码仅作参考）——

	'*********************************************************************
else '验证失败
    '如要调试，请看alipay_notify.asp页面的VerifyReturn函数，比对sign和mysign的值是否相等，或者检查responseTxt有没有返回true
    	Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""xwbox""><p>您的订单，出了问题！</p></div>"
		Response.Write"<div class=""body""><span>提示：</span>您的支付订单出了问题，如果您确实已经付款完成，请到支付宝找到交易订单号联系本站管理员！</div>"
		Response.Write"</div>"
end if%>
</body>
</html>
