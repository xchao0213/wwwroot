<!--#include file="../../Include/conn.asp" -->
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<%
' ���ܣ�֧����ҳ����תͬ��֪ͨҳ��
' �汾��3.3
' ���ڣ�2012-07-17
' ˵����
' ���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
' �ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���
	
' //////////////ҳ�湦��˵��//////////////
' ��ҳ����ڱ������Բ���
' �ɷ���HTML������ҳ��Ĵ��롢�̻�ҵ���߼��������
' ��ҳ�����ʹ��ASP�������ߵ��ԣ�Ҳ����ʹ��д�ı�����LogResult���е��ԣ��ú����ѱ�Ĭ�Ϲرգ���alipay_notify.asp�еĺ���VerifyReturn
' ���ýӿڵ�ע��������û����ô����ɾ������
'////////////////////////////////////////
%>
<!--#include file="class/alipay_notify.asp"-->
<title>֧����֧�����</title>
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
'����ó�֪ͨ��֤���
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyReturn()

If sVerifyResult Then'��֤�ɹ�
	'*********************************************************************
	'������������̻���ҵ���߼��������
	'�������������ҵ���߼�����д�������´�������ο�������
    '��ȡ֧������֪ͨ���ز������ɲο������ĵ���ҳ����תͬ��֪ͨ�����б�
	'�̻�������
	out_trade_no = Request.QueryString("out_trade_no")

	'֧�������׺�
	trade_no = Request.QueryString("trade_no")
	'����״̬
	trade_status = Request.QueryString("trade_status")

	If Request.QueryString("trade_status") ="WAIT_SELLER_SEND_GOODS" Then '����Ѹ���ȴ����ҷ���
	    '�ж��Ƿ����̻���վ���Ѿ����������֪ͨ���صĴ���
		'���û������������ôִ���̻���ҵ�����
		'���������������ô��ִ���̻���ҵ�����
		conn.execute("update orders set state='2' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		SendEmail ""&email&"","���ã���������:"&title&"","���ã����Ѿ����˱�վ��Ʒ��"&title&"�����ǽ�������㷢����","Jmail" 

	ElseIf request.QueryString("trade_status") = "TRADE_FINISHED" Then '��ɽ���
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
		'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
		'���������������ִ���̻���ҵ�����
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
		body=body+"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""700"" height=""39"" style=""font:12px Tahoma, Arial, ����;"">"
		body=body+"<tbody><tr>"
		body=body+"<td width=""210""><a target=_blank href="&zych_url&"><img src="""&zych_url&"/images/logo.png"""" height=""39"" border=""0""></a></td>"
		body=body+"<td width=""490"" align=""right"" valign=""bottom"" style=""padding-bottom:10px;""><a target=""_blank"" style=""color:#07f;text-decoration:none;"" href="""&zych_url&"/Member/"">��Ա����</a></td>"
		body=body+"</tr></tbody></table></div>"
		body=body+"<div style=""width:680px;padding:0 10px;"">"
		body=body+"<div style=""line-height:1.5;font-size:14px;margin-bottom:25px;color:#4d4d4d;""><strong style=""display:block;margin-bottom:15px;"">�װ��Ļ�Ա���ã�</strong>"
		body=body+"<p>�������� "&Request.QueryString("notify_time")&" �ɹ��� "&zych_url&" ������ ��<strong style=""font-size:14px;""><a target=_blank style=color:#f60; href="&zych_url&dir&"Products/ShowProducts.asp?id="&proid&">"&title&"</a></strong></p></div>"
		body=body+"<div style=""margin-bottom:30px;"">"
		body=body+"<p>��Ʒ������<span>"&Request.QueryString("out_trade_no")&"</span></p>"
		body=body+"<p>�����<span>"&Request.QueryString("total_fee")&"Ԫ</span></p>"
		body=body+"<p>�ʼķ��ã�<span>"&Request.QueryString("logistics_fee")&"Ԫ</span></p>"
		body=body+"<p>����״̬��<span>"&Request.QueryString("trade_status")&" (�������)</span></p>"
		body=body+"<p>֧�������׺ţ�<span>"&Request.QueryString("trade_no")&"</span></p>"
		body=body+"<p>������ڽ��׹���������ʲô���������Ϣ����Ϊ����ƾ֤��</p>"
		body=body+"<a href="&zych_url&dir&"Products/ShowProducts.asp?id="&proid&" target=_blank>"&zych_url&dir&"Products/ShowProducts.asp?id="&proid&"</a></div>"
		body=body+"<div style=""margin-bottom:30px;"">"
		body=body+"<p>"&zych_description&"</p>"
		body=body+"<p><a href="&zych_url&" target=_blank>"&zych_url&"</a></p>"
		body=body+"</div></div>"
		body=body+"<div style=""padding:10px 10px 0;border-top:1px solid #ccc;color:#999;margin-bottom:20px;line-height:1.3em;font-size:12px;"">"
		body=body+"<p style=""margin-bottom:15px;"">��Ϊϵͳ�ʼ�������ظ�<br>�뱣�ܺ��������䣬����֧�����˻������˵���<br>Copyright zych 2010-2014 All Right Reserved</p>"
		body=body+"</div>  </div>"
		End Function
		SendEmail ""&email&"","���Ѿ�������:"&title&"",""&body&"","Jmail" 
	Else
		Response.Write "trade_status="&Request.QueryString("trade_status")
	End If
		Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""cgbox""><p>���Ķ�����֧���ɹ��ˣ�</p></div>"
		Response.Write"<div class=""body""><span>��ʾ��</span>���Ѿ��ɹ�����,���ڴ������Ķ�������رմ���ҳ�档��������Զ���ת��</div>"
		Response.Write"</div>"
	'�������������ҵ���߼�����д�������ϴ�������ο�������

	'*********************************************************************
else '��֤ʧ��
    '��Ҫ���ԣ��뿴alipay_notify.aspҳ���VerifyReturn�������ȶ�sign��mysign��ֵ�Ƿ���ȣ����߼��responseTxt��û�з���true
    	Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""xwbox""><p>���Ķ������������⣡</p></div>"
		Response.Write"<div class=""body""><span>��ʾ��</span>����֧�������������⣬�����ȷʵ�Ѿ�������ɣ��뵽֧�����ҵ����׶�������ϵ��վ����Ա��</div>"
		Response.Write"</div>"
end if%>
</body>
</html>
