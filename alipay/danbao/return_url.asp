<!--#include file="../../Include/conn.asp" -->
<%
' ���ܣ�֧����ҳ����תͬ��֪ͨҳ��' �汾��3.3' ���ڣ�2012-07-17%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
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

If sVerifyResult Then	'��֤�ɹ�
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

	
	If Request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" Then
	'�ж��Ƿ����̻���վ���Ѿ����������֪ͨ���صĴ���
		'���û������������ôִ���̻���ҵ�����
		'���������������ô��ִ���̻���ҵ�����
		conn.execute("update orders set state='2' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		
	ElseIf request.QueryString("trade_status") = "TRADE_FINISHED" Then '��ɽ���
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
		'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
		'���������������ִ���̻���ҵ�����
		conn.execute("update orders set state='8' where OrderNo='"&out_trade_no&"'")
		conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
	Else
		Response.Write "trade_status="&Request.QueryString("trade_status")
	End If

	'Response.Write "��֤�ɹ�<br>"
		Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""cgbox""><p>���Ķ�����֧���ɹ��ˣ�</p></div>"
		Response.Write"<div class=""body""><span>��ʾ��</span>���Ѿ��ɹ�����,���ڴ������Ķ�������رմ���ҳ�档��������Զ���ת��</div>"
		Response.Write"</div>"
	'�������������ҵ���߼�����д�������ϴ�������ο�������
	
	'*********************************************************************
else '��֤ʧ��
    '��Ҫ���ԣ��뿴alipay_notify.aspҳ���VerifyReturn�������ȶ�sign��mysign��ֵ�Ƿ���ȣ����߼��responseTxt��û�з���true
    response.Write "��֤ʧ��"
end if
%>
</body>
</html>
