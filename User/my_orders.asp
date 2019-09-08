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
<title>会员中心_<%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<link rel="stylesheet" type="text/css" href="../css/style.css"/>
<script type="text/javascript" src="<%=dir%>js/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="<%=dir%>js/dropdown.js"></script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>会员中心 <i>User</i></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>User/">会员中心</a>我的订单</span></div>
<div class="w_650 fl">
<div class="position">
<table width="100%" border="1" cellpadding="10" cellspacing="0" bordercolor="#EDEDED" Class="fkzt yh">
  <tr>
    <td width="25" align="center">1.等待买家付款</td>
    <td width="25" align="center">2.已付款,等待发货</td>
    <td width="25" align="center">3.已发货,等待确认收货</td>
    <td width="25" align="center">4.确认收货,交易完成</td>
  </tr>
</table>
<div class="ubox">
<div class="pronav"><h1>最近订单</h1></div>


<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#EDEDED" style="border-collapse: collapse;line-height:25px; text-indent:3px">
<tr>
    <td width="19%" align="center"><strong>订单编号</strong></td>
    <td width="27%" align="center"><strong>产品名称</strong></td>
    <td width="23%" align="center"><strong>时 间</strong></td>
    <td width="19%" align="center"><strong>状 态</strong></td>
    <td colspan="2" align="center"><strong>操 作</strong></td>
    </tr>
<%set rs=server.createobject("adodb.recordset") 
exec="select * from [orders] where userid="&session("username")&" order by id desc "  
rs.open exec,conn,1,1 
if rs.eof then
response.Write "&nbsp;暂无订单！"
else
rs.PageSize =15 '每页记录条数
iCount=rs.RecordCount '记录总数
iPageSize=rs.PageSize
maxpage=rs.PageCount 
page=request("page")
if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
if page=maxpage then
x=iCount-(maxpage-1)*iPageSize
else
x=iPageSize
end if	
For i=1 To x
if rs("state")=1 then
ddzt="<a href="""&dir&"alipay/"&alipayLX&"/alipayapi_re.asp?id="&rs("id")&""" title=""点击这里完成付款"" class=""kfal_a"" target=""_blank"" class=""kfal_a"" onClick=""document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'""><font>立即付款</font></a>"
elseif rs("state")=2 then
ddzt="买家已付款，等待发货"
elseif rs("state")=3 then
ddzt="已发货，等待确认收货"
elseif rs("state")=8 then
ddzt="交易完成"
else
ddzt="订单出现异常"
end if%>
  <tr>
    <td width="19%" align="center"><a href="my_Showorders.asp?id=<%=rs("id")%>" title="点击查看订单详情"><%=rs("OrderNo")%></a></td>
    <td width="27%"><%=stvalue(rs("title"),24)%></td>
    <td width="23%" align="center"><%=rs("data")%></td>
    <td width="19%" align="center"><%=ddzt%></td>
    <td width="6%" align="center"><a href="my_Showorders.asp?id=<%=rs("id")%>">查看</a></td>
    <td width="6%" align="center"><a href="javascript:" onClick="javascript:if(confirm('确定删除该订单？删除后不可恢复!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}" title="删除此订单">删除</a></td>
  </tr>
<%rs.movenext
next
end if%>
</table>
<%'以下显示分页
call PageControl(iCount,maxpage,page)
rs.close
set rs=nothing%>
</div></div>
</div>
   
<div class="w_280 fr">
    <div class="box">
      <div class="title tubiao_U">会员导航 <span>User Nav</span></div>
      <ul class="list_2">
        <%=zych_user_nav%>
      </ul>
    </div>
    <!--#include file="../Include/left.asp" -->
  </div>
</div>
<div id="light" class="Order_tc">
<span class="ortitle">网上支付提示</span>
<div class="up-box">
<span class="icon"><img src="../images/working.gif"></span>
<div class="r-txt">支付完成前，请不要关闭此支付验证窗口。<br clear="none">支付完成后，请根据您支付的情况点击下面按钮。</div>
</div> 
<div class="lay-btn">
<a href="/User" class="btn92" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">支付遇到问题</a>
<a href="/User/my_orders.asp" class="btn92s" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'" shape="rect">支付完成</a>
</div>
</div> 
<div id="fade" class="Order_bj"> </div>
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
Response.Write "<script>alert('当前订单删除成功！');window.location.href='my_orders.asp';</script>"
end if %>