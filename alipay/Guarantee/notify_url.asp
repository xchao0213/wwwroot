<!--#include file="../../Include/conn.asp" -->
<%
' ���ܣ�֧�����������첽֪ͨҳ��
' �汾��3.3
' ���ڣ�2012-07-17
' ˵����
' ���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
' �ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���
	
' //////////////ҳ�湦��˵��//////////////
' ������ҳ���ļ�ʱ�������ĸ�ҳ���ļ������κ�HTML���뼰�ո�
' ��ҳ�治���ڱ������Բ��ԣ��뵽�������������ԡ���ȷ���ⲿ���Է��ʸ�ҳ�档
' ��ҳ����Թ�����ʹ��д�ı�����LogResult���ú����ѱ�Ĭ�Ͽ�������alipay_notify.asp�еĺ���VerifyNotify
' ���û���յ���ҳ�淵�ص� success ��Ϣ��֧��������24Сʱ�ڰ�һ����ʱ������ط�֪ͨ
' 
' ����жϸñʽ�����ͨ����ʱ���ʷ�ʽ�����ͨ���������׷�ʽ���
'
'�������׵Ľ���״̬�仯˳���ǣ��ȴ���Ҹ��������Ѹ���ȴ����ҷ����������ѷ������ȴ�����ջ���������ջ����������
'��ʱ���ʵĽ���״̬�仯˳���ǣ��ȴ���Ҹ�����������
'
'ÿ���յ�֧��������֪ͨʱ���Ϳ��Ի�ȡ����ʽ��׵Ľ���״̬�������̻���Ҫ�����̻������Ų�ѯ�̻���վ�Ķ������ݣ�
'�õ���ʶ������̻���վ�е�״̬��ʲô�����̻���վ�еĶ���״̬���֧����֪ͨ�л�ȡ����״̬�����Աȡ�
'����̻���վ��Ŀǰ��״̬�ǵȴ���Ҹ������֧����֪ͨ��ȡ����״̬������Ѹ���ȴ����ҷ�������ô��ʽ���������õ������׷�ʽ�����
'����̻���վ��Ŀǰ��״̬�ǵȴ���Ҹ������֧����֪ͨ��ȡ����״̬�ǽ�����ɣ���ô��ʽ���������ü�ʱ���ʷ�ʽ�����
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
'����ó�֪ͨ��֤���
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyNotify()

if sVerifyResult then	'��֤�ɹ�
	'*********************************************************************
	'������������̻���ҵ���߼��������
	'�������������ҵ���߼�����д�������´�������ο�������
    '��ȡ֧������֪ͨ���ز������ɲο������ĵ��з������첽֪ͨ�����б�
	'�̻�������
	out_trade_no = Request.Form("out_trade_no")
	'֧�������׺�
	trade_no = Request.Form("trade_no")
	'����״̬
	trade_status = Request.Form("trade_status")
	set rs=server.createobject("adodb.recordset") 
	exec="select * from [orders] where OrderNo='"&out_trade_no&"'"
	rs.open exec,conn,1,1 
	if rs.eof and rs.bof then
		Response.Write"�޴˶���"
	else
		state=rs("state")
	end if
	rs.close
	set rs=nothing
	
	Private Sub Handle(zt)
		set db=conn.execute("select * from [orders] where OrderNo='"&out_trade_no&"'")
		if db.eof and db.bof then
		   Response.Write"�޴˶���"
		else
		  if db("state")<>zt then
		  	conn.execute("update orders set state='"&zt&"' where OrderNo='"&out_trade_no&"'")
			conn.execute("update orders set alipayno='"&trade_no&"' where OrderNo='"&out_trade_no&"'")
		  else
		  	Response.Write "success"
		  end if
	    end if
	End Sub	
	
	
	If Request.Form("trade_status") = "WAIT_BUYER_PAY" and state=1 Then '�ȴ���Ҹ���
	
	    Call Handle(1)
     	'���жϱ�ʾ�������֧�������׹����в����˽��׼�¼����û�и���
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
		'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
		'���������������ִ���̻���ҵ�����
		Response.Write "success"	'�벻Ҫ�޸Ļ�ɾ��

		'�����ã�д�ı�������¼������������Ƿ�����
        'LogResult("����д����Ҫ���ԵĴ������ֵ�����������еĽ����¼")
	Elseif Request.Form("trade_status") = "WAIT_SELLER_SEND_GOODS" Then '����Ѹ���ȴ����ҷ���
		'Call Handle(2)
	    '���жϱ�ʾ�������֧�������׹����в����˽��׼�¼�Ҹ���ɹ���������û�з���
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
		'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
		'���������������ִ���̻���ҵ�����
		Response.Write "success"	'�벻Ҫ�޸Ļ�ɾ��
		
		'�����ã�д�ı�������¼������������Ƿ�����
        'LogResult("����д����Ҫ���ԵĴ������ֵ�����������еĽ����¼")
	Elseif Request.Form("trade_status") = "WAIT_BUYER_CONFIRM_GOODS" and state<>3 Then '�����ѷ������ȴ�����ջ�
	   Call Handle(3)
		'���жϱ�ʾ�����Ѿ����˻�������һ�û����ȷ���ջ��Ĳ���
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
		'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
		'���������������ִ���̻���ҵ�����
		Response.Write "success"	'�벻Ҫ�޸Ļ�ɾ��
		'�����ã�д�ı�������¼������������Ƿ�����
        'LogResult("����д����Ҫ���ԵĴ������ֵ�����������еĽ����¼")
	Elseif Request.Form("trade_status") = "TRADE_FINISHED" and state<>8 Then '������ջ����������
		 Call Handle(8)
	    '���жϱ�ʾ����Ѿ�ȷ���ջ�����ʽ������
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
		'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
		'���������������ִ���̻���ҵ�����
		Response.Write "success"	'�벻Ҫ�޸Ļ�ɾ��
		'�����ã�д�ı�������¼������������Ƿ�����
        'LogResult("����д����Ҫ���ԵĴ������ֵ�����������еĽ����¼")
	Else
		'����״̬�жϡ�
		'Call Handle(4)
		Response.Write "success"
		'�����ã�д�ı�������¼������������Ƿ�����
		'LogResult ("����д����Ҫ���ԵĴ������ֵ�����������еĽ����¼")
	End If
		Response.Write"<div class=""mbox"">"
		Response.Write"<div class=""cgbox""><p>���Ķ�����֧���ɹ��ˣ�</p></div>"
		Response.Write"<div class=""body""><span>��ʾ��</span>���Ѿ��ɹ�����,���ڴ������Ķ�������رմ���ҳ�档��������Զ���ת��</div>"
		Response.Write"</div>"

	'�������������ҵ���߼�����д�������ϴ�������ο�������
	'*********************************************************************
else '��֤ʧ��
    response.Write "��Ķ����������⣨fail��"
	'�����ã�д�ı�������¼������������Ƿ�����
	'LogResult("����д����Ҫ���ԵĴ������ֵ�����������еĽ����¼")
	Response.Write"<div class=""mbox"">"
	Response.Write"<div class=""xwbox""><p>���Ķ������������⣡</p></div>"
	Response.Write"<div class=""body""><span>��ʾ��</span>����֧�������������⣬�����ȷʵ�Ѿ�������ɣ��뵽֧�����ҵ����׶�������ϵ��վ����Ա��</div>"
	Response.Write"</div>"
end if
%>