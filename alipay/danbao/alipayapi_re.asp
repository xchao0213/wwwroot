<!--#include file="../../Include/conn.asp" -->
<!--#include file="class/alipay_submit.asp" -->
<%id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [orders] where id="&id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title>֧��������_<%=zych_home%></title>
<link href="../images/style.css" rel="stylesheet" type="text/css" />
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

'����֧�����ʻ�
seller_email=alipayID
'����
out_trade_no =rs("OrderNo") '�̻�������,����ϵͳ��Ψһ�����ţ�����
cpid=rs("cpid")'��ƷID
userid=rs("userid")'��վ�û�ID
subject =rs("title")'��������
price = rs("rmb")'������
quantity =rs("number")'��Ʒ����'�������Ĭ��Ϊ1�����ı�ֵ����һ�ν��׿�����һ���¶������ǹ���һ����Ʒ
logistics_fee =rs("kuaidi")'��������
logistics_type = "EXPRESS"'��������
'�������ֵ��ѡ��EXPRESS����ݣ���POST��ƽ�ʣ���EMS��EMS��
'����֧����ʽ
logistics_payment = "BUYER_PAY"
'�������ֵ��ѡ��SELLER_PAY�����ҳе��˷ѣ���BUYER_PAY����ҳе��˷ѣ�

show_url =""&zych_url&"Products/ShowProducts.asp?id="&PID&""'��Ʒչʾ��ַ
'����http://��ͷ������·�����磺http://www.xxx.com/myorder.html

receive_name =rs("name")'�ջ�������
qq=rs("qq")
quantity=rs("number")
email=rs("email")

receive_address =rs("address")'�ջ��˵�ַ
'�磺XXʡXXX��XXX��XXX·XXXС��XXX��XXX��ԪXXX��

'�ջ��˵绰����
receive_phone =rs("tel")
'�磺0571-88158090

'�ջ����ֻ�����
receive_mobile =rs("tel")
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
