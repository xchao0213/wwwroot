<!--#include file="../../Include/conn.asp" -->
<!--#include file="class/alipay_submit.asp"-->
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title>֧��������_<%=zych_home%></title>
<link href="../images/main.css" rel="stylesheet" type="text/css" />
</head>
<body>

<%
'/////////////////////�������/////////////////////

'֧������
payment_type = "1"
'��������޸�
'�������첽֪ͨҳ��·��
notify_url = ""&zych_url&"/alipay/danbao/notify_url.asp"
'��http://��ʽ������·�������ܼ�?id=123�����Զ������

'ҳ����תͬ��֪ͨҳ��·��
return_url = ""&zych_url&"/alipay/danbao/return_url.asp"
'��http://��ʽ������·�������ܼ�?id=123�����Զ������������д��http://localhost/

seller_email=alipayID'����
'�̻�������
out_trade_no = Request.Form("out_trade_no")
'�̻���վ����ϵͳ��Ψһ�����ţ�����
userid=Request.Form("userid")
'��������
subject = Request.Form("WIDtitle")
'����
cpid=Request.Form("cpid")
'������
price = cint(Request.Form("WIDprice"))+cint(Request.Form("WIDtax"))
'����

'��Ʒ����
quantity = Request.Form("quantity")
'�������Ĭ��Ϊ1�����ı�ֵ����һ�ν��׿�����һ���¶������ǹ���һ����Ʒ
'��������
logistics_fee = Request.Form("WIDFreight")
'������˷�
'��������
logistics_type = "EXPRESS"
'�������ֵ��ѡ��EXPRESS����ݣ���POST��ƽ�ʣ���EMS��EMS��
'����֧����ʽ
logistics_payment = "SELLER_PAY"
'�������ֵ��ѡ��SELLER_PAY�����ҳе��˷ѣ���BUYER_PAY����ҳе��˷ѣ�
'��������

body = Request.Form("WIDbody")
'��Ʒչʾ��ַ
show_url = Request.Form("WIDshow_url")
'����http://��ͷ������·�����磺http://www.�̻���վ.com/myorder.html

'�ջ�������
receive_name = Request.Form("WID_name")
'�磺����
qq= Request.Form("qq")
web= Request.Form("WIDweb")
email= Request.Form("email")
'�ջ��˵�ַ
receive_address = Request.Form("WIDreceive_address")
'�磺XXʡXXX��XXX��XXX·XXXС��XXX��XXX��ԪXXX��

'�ջ����ʱ�
receive_zip = Request.Form("WIDreceive_zip")
'�磺123456

'�ջ��˵绰����
receive_phone = Request.Form("WIDreceive_phone")
'�磺0571-88158090

'�ջ����ֻ�����
receive_mobile = Request.Form("WID_tel")
'�磺13312341234

'/////////////////////�������/////////////////////

'���������������
sParaTemp = Array("service=create_partner_trade_by_buyer","partner="&partner,"seller_email="&seller_email,"_input_charset="&input_charset  ,"payment_type="&payment_type   ,"notify_url="&notify_url   ,"return_url="&return_url   ,"out_trade_no="&out_trade_no   ,"subject="&subject   ,"price="&price   ,"quantity="&quantity   ,"logistics_fee="&logistics_fee   ,"logistics_type="&logistics_type   ,"logistics_payment="&logistics_payment   ,"body="&body   ,"show_url="&show_url   ,"receive_name="&receive_name   ,"receive_address="&receive_address   ,"receive_zip="&receive_zip   ,"receive_phone="&receive_phone   ,"receive_mobile="&receive_mobile  )

'��������
Set objSubmit = New AlipaySubmit
sHtml = objSubmit.BuildRequestForm(sParaTemp, "get", "ȷ��")
response.Write sHtml


%>
</body>
</html>
