<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!--#include file="../Include/seeion.asp" -->
<!--#include file="../Include/page.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ա����_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
<script language="javascript">
$(document).ready(function()
{
$("#btnSnap").click(function()
{
    $("#retData").html('loading...');

    var expressid = $("#expressid").val();
    var expressno = $("#expressno").val();
    $.get("express.asp",{com:expressid,nu:expressno},
        function(data)
        {
            $("#retData").html(data);
        }
    );

    return false;
});
});
</script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>��Ա���� <i>User</i></b><span>��ǰλ�ã�<a href="../">��ҳ</a><a href="<%=dir%>User/">��Ա����</a>�ҵĶ���</span></div>
<div class="w_650 fl">
<div class="position">
<%id=request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('��������');history.go(-1);</script>" 
Response.End()
end if
exec="select * from [orders] where id="& id 
set rs=server.createobject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
Response.Write "<script>alert('��������ȷ��IDֵ�����ڣ�');history.go(-1);</script>" 
Response.End()
end if%>
<table width="100%" border="0" cellpadding="10" cellspacing="0" bordercolor="#EDEDED" Class="fkzt yh">
  <tr>
    <td width="25" align="center" <% If rs("state")=1 Then %>class="zt" <% ElseIf rs("state")>1 Then %>class="zg"<%End If%>>�ȴ���Ҹ���</td>
    <td width="25" align="center" <% If rs("state")=2 Then %>class="zt" <% ElseIf rs("state")>2 Then %>class="zg"<%End If%>>�Ѹ���,�ȴ�����</td>
    <td width="25" align="center" <% If rs("state")=3 Then %>class="zt" <% ElseIf rs("state")>3 Then %>class="zg"<%End If%>>�ѷ���,�ȴ�ȷ���ջ�</td>
    <td width="25" align="center" <% If rs("state")=8 Then %>class="zt"<%End If%>>ȷ���ջ�,�������</td>
  </tr>
</table>

<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse:collapse;line-height:25px; text-indent:3px">
 <tr>
  <td align="right">������ţ�</td>
  <td width="88%" colspan="2"><%=rs("OrderNo")%></td>
  </tr>
   <tr>
  <td align="right">������Ʒ��</td>
  <td width="88%" colspan="2">
<div style="width:auto; position:relative; display:block">
<%set rsp=server.createobject("adodb.recordset") 
exec="select * from [Products] where id in ("&rs("cpid")&")"
rsp.open exec,conn,1,1
if rsp.eof then
response.Write "<div style=""padding:10px"">�޴���Ʒ</div>"
response.End()
end if
while not rsp.eof%>
<div style=" width:110px; height:80px; float:left;"><img src="<%=rsp("img")%>" width="100" height="60" alt="<%=rsp("title")%>"/><p><%=stvalue(rsp("title"),16)%></p></div>
<%rsp.movenext
wend
rsp.close
set rsp=nothing %>
</div> 
  </td>
  </tr>
<tr>
  <tr>
  <td align="right">�����</td>
  <td width="88%" colspan="2"><font style="font-size:18px; color:#F00">��<%=rs("rmb")%></font> (��Ʒ���) 
  <font style="font-size:16px; color:#F00">��<%=rs("kuaidi")%></font> {�ʼķ���}</td>
 </tr>
<tr>
  <td align="right">����������</td>
  <td width="88%" colspan="2"><%=rs("name")%></td>
 </tr>
<tr>
  <td align="right">��ϵ�绰��</td>
  <td colspan="2"><%=rs("tel")%></td>
  </tr>
<tr>
  <td align="right">��ϵ��ַ��</td>
  <td colspan="2"><%=rs("address")%></td>
  </tr>
<tr>
  <td align="right">��ϵQQ��</td>
  <td colspan="2"><%=rs("QQ")%></td>
</tr>
<tr>
  <td align="right">����˵����</td>
  <td colspan="2"><%=rs("sm")%></td>
</tr>
<tr>
  <td align="right">����ʱ�䣺</td>
  <td colspan="2"><%=rs("data")%></td>
</tr>
<%if rs("state")=1 then
ddzt="<a href="""&dir&"alipay/"&alipayLX&"/alipayapi_re.asp?id="&rs("id")&""" title=""���������ɸ���"" class=""kfal_a""><font>��������</font></a>"
elseif rs("state")=2 then
ddzt="����Ѹ���ȴ�����"
elseif rs("state")=3 then
ddzt="�ѷ������ȴ�ȷ���ջ�"
elseif rs("state")=8 then
ddzt="�������"
else
ddzt="���������쳣"
end if%>
 <tr>
  <td align="right">����״̬��</td>
  <td colspan="2"><%=ddzt%></td>
</tr>
<%state=rs("state") 
if state=2 then
response.Write("<tr>")
response.Write("<td>���������</td>")
response.Write("<td colspan=""2""><div style=""text-align:center;line-height:35px"">���Ķ�����δ(����)������������<a href="""&dir&"alipay/"&alipayLX&"/alipayapi_re.asp?id="&rs("id")&""" title=""����������¸���""><font style=color:#FF0000;font-size:14px;>�������</font></a>��ɸ���!</div></td>")
response.Write(" </tr>")
else%>
<tr>
<td align="right">��ע��Ϣ��</td>
<td colspan="2"><font color="#FF0000"><%=rs("bz")%></font></td>
</tr>
<%end if%>
</table>
<% if rs("state")<>1 then%>
<div id="retForm">
<div class="txtURL">
<div class="hdtxt">��ݹ�˾��
<%set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from orders_kd where kddm='"&rs("kddm")&"'",conn,1,1
if rsc.eof and rsc.bof then
Response.Write("��δ�ʼ�")
else
response.Write("<a href="""&rsc("url")&"""><font color=#003399>["&rsc("title")&"]</font></a>")
end if
rsc.close
set rsc=nothing%>
<input name="expressid" type="hidden" id="expressid" value="<%=rs("kddm")%>"/></div>
<div class="dhtxt">��ݵ���</div>
<input name="expressno" class="dh" type="text" id="expressno" placeholder="�������ݵ���" value="<%=rs("kddh")%>"/>
<input type="submit" value="��ѯ" id="btnSnap" class="btn"/>
</div>
</div>
<div class="txtAboutSnap">
<div id="retData"></div>
</div>
<%end if%>
</div>
</div>
   
<div class="w_280 fr">
    <div class="box">
      <div class="title tubiao_U">��Ա���� <span>User Nav</span></div>
      <ul class="list_2">
        <%=zych_user_nav%>
      </ul>
    </div>
    <!--#include file="../Include/left.asp" -->
  </div>
</div>

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<%if request("del")="ok" then
set rs=server.createobject("adodb.recordset")
id=Request.QueryString("id")
sql="select * from orders where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('��ǰ����ɾ���ɹ���');window.location.href='my_orders.asp';</script>"
end if %>