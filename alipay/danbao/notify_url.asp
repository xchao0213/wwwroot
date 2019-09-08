<!--#include file="../../Include/conn.asp" -->
<%
' 功能：支付宝服务器异步通知页面' 版本：3.3' 日期：2012-07-17
%>
<style>
.mbox{width:600px; height:200px; position:absolute; left:50%; top:50%; margin-left:-300px; margin-top:-100px;background:#edffcd; border:1px #9bdb51 solid}
.mbox .cgbox{ height:80px;border-bottom:1px #CCCCCC solid}
.mbox .cgbox p{width:380px; padding-left:50px; font-size:30px;line-height:42px; margin:40px auto 20px auto; background:url(../images/gxico.png) no-repeat;}
.mbox .xwbox{ height:80px;border-bottom:1px #CCCCCC solid}
.mbox .xwbox p{width:380px; padding-left:50px; font-size:30px;line-height:42px; margin:40px auto 20px auto; background:url(../images/gxico.png) 0px -42px no-repeat;}
.mbox .body{ padding:10px; text-align:center; font-size:12px;}
.mbox .body span{font-weight:bold; padding-right:5px}
</style>
<!--#include file="class/alipay_notify.asp"-->

<%
'计算得出通知验证结果
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyNotify()

if sVerifyResult then	'验证成功
	'*********************************************************************
	'请在这里加上商户的业务逻辑程序代码

	'——请根据您的业务逻辑来编写程序（以下代码仅作参考）——
    '获取支付宝的通知返回参数，可参考技术文档中服务器异步通知参数列表

	'商户订单号

	out_trade_no = Request.Form("out_trade_no")

	'支付宝交易号

	trade_no = Request.Form("trade_no")

	'交易状态
	trade_status = Request.Form("trade_status")
	set rs=server.createobject("adodb.recordset") 
	exec="select * from [orders] where OrderNo='"&out_trade_no&"'"
	rs.open exec,conn,1,1 
	if rs.eof and rs.bof then
		Response.Write"无此订单"
	else
		state=rs("state")
	end if
	rs.close
	set rs=nothing
	
	Private Sub Handle(zt)
		set db=conn.execute("select * from [orders] where OrderNo='"&out_trade_no&"'")
		if db.eof and db.bof then
		   Response.Write"无此订单"
		else
		  if db("state")<>zt then
		  	conn.execute("update orders set state='"&zt&"' where OrderNo='"&out_trade_no&"'")
			conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		  else
		  	Response.Write "success"
		  end if
	    end if
	End Sub	

	
	If Request.Form("trade_status") = "WAIT_BUYER_PAY" and state=1 Then
	'该判断表示买家已在支付宝交易管理中产生了交易记录，但没有付款
		Call Handle(1)
		'判断该笔订单是否在商户网站中已经做过处理
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		Response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'LogResult("这里写入想要调试的代码变量值，或其他运行的结果记录")
	Elseif Request.Form("trade_status") = "WAIT_SELLER_SEND_GOODS" Then
	'该判断表示买家已在支付宝交易管理中产生了交易记录且付款成功，但卖家没有发货
		Call Handle(2)
		'判断该笔订单是否在商户网站中已经做过处理
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		Response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'LogResult("这里写入想要调试的代码变量值，或其他运行的结果记录")
	Elseif Request.Form("trade_status") = "WAIT_BUYER_CONFIRM_GOODS" and state<>3  Then
	'该判断表示卖家已经发了货，但买家还没有做确认收货的操作
	Call Handle(3)
		'判断该笔订单是否在商户网站中已经做过处理
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		Response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'LogResult("这里写入想要调试的代码变量值，或其他运行的结果记录")
	Elseif Request.Form("trade_status") = "TRADE_FINISHED"   and state<>8 Then
	'该判断表示买家已经确认收货，这笔交易完成
	 Call Handle(8)
		'判断该笔订单是否在商户网站中已经做过处理
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		Response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'LogResult("这里写入想要调试的代码变量值，或其他运行的结果记录")
	Else
		'其他状态判断。
		
		Response.Write "success"
		'调试用，写文本函数记录程序运行情况是否正常
		'LogResult ("这里写入想要调试的代码变量值，或其他运行的结果记录")
	End If
		Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""cgbox""><p>您的订单，支付成功了！</p></div>"
		Response.Write"<div class=""body""><span>提示：</span>您已经成功付款,正在处理您的订单请勿关闭此自页面。处理完毕自动跳转！</div>"
		Response.Write"</div>"
	'——请根据您的业务逻辑来编写程序（以上代码仅作参考）——
	
	'*********************************************************************
else '验证失败
    response.Write "fail"
	'调试用，写文本函数记录程序运行情况是否正常
	'LogResult("这里写入想要调试的代码变量值，或其他运行的结果记录")
end if
%>