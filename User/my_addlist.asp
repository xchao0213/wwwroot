<!--#include file="../Include/conn.asp" -->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"--> 
<!--#include file="../Include/md5.Asp" -->
<!--#include file="../Include/seeion.asp" -->
<!--#include file="../Include/page.asp" -->
<%randNumber=getStrRandNumber(1000,9999)
BuyId=getTime&randNumber%>
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
<script type="text/javascript">          
function TopicShow(e,TopicID){
e.className=(e.className=="btn1")?"btn2":"btn1"
document.getElementById(TopicID).style.display=(e.className=="btn1")?"":"none"
}//显示隐藏层
document.myform.k_hd.value='#f1f1f1';
</script>
</head>
<body>
<!--#include file="../Include/top.asp" -->
<div class="main">
<div class="ad_flash"><%=zych_Flash()%></div>
<div class="bar"><b>会员中心 <i>User</i></b><span>当前位置：<a href="../">首页</a><a href="<%=dir%>User/">会员中心</a>我的购物车</span></div>
<div class="w_650 fl">
<div class="position">
<div class="ubox">
<div class="pronav"><h1>我的购物车</h1></div>
<table width="100%"  border="1" cellpadding="6" cellspacing="0" bordercolor="#dedede" style="border-collapse:collapse;line-height:22px; text-indent:3px">
  <tr><td colspan="8" align="center" height="40px"><b style="font-size:24px">我的购物车</b></td></tr>
  <tr style="background: #E7E7E7; font-weight:bold">
    <td width="42"  align="center">序 号</td>
    <td colspan="2"  align="center">商品名称</td>
    <td width="63" align="center">数 量</td>
    <td width="52" align="center">单 价</td>
    <td width="52" align="center">金 额</td>
    <td colspan="2" align="center">操 作</td>
  </tr>
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [ogwc] where sfdd=0 and userid="&session("username")&" order by id desc",conn,1,1
if rs.eof and rs.bof then
response.Write("<tr><td colspan=""7"" align=""center"">暂无进货商品...</td></tr>")
end if
while not rs.eof
i=i+1
set rsp=server.createobject("adodb.recordset") 
exec="select * from [Products] where id="&rs("cpid")&""
rsp.open exec,conn,1,1 
if rsp.eof and rsp.bof then
response.Write""
else
cp_title=rsp("title")
cp_zkj=rsp("zkj")
cp_img=rsp("img")
cp_rmb=rsp("zkj")*rs("cpsl")
end if
rsp.close
set rsp=nothing%>
<form action="?id=<%=rs("id")%>&xiugai=ok" method="post" name="add">
<tr>
  <td align="center"><%=right("00"&i,3)%><input name="id" type="hidden" size="5"  value="<%=rs("id")%>"/></td>
  <td width="98" align="center"><img src="<%=cp_img%>" alt="<%=cp_title%>" width="100"></td>
  <td width="152"><%=cp_title%></td>
  <td align="center"><input name="cpsl" type="text" style="width:48px;text-align:center" value="<%=rs("cpsl")%>" /></td>
  <td align="center"><%=cp_zkj%><input name="zkj" type="hidden" size="5"  value="<%=cp_zkj%>"/></td>
  <td align="center"><%=cp_rmb%></td>
  <td width="42" align="center"><input type="submit" name="button" value="修改"/></td>
  <td width="47" align="center"><input type="button" name="button2" value="删除" onClick="javascript:if(confirm('确定删除该订单？删除后不可恢复!')){window.location.href='?id=<%=rs("id")%>&amp;del=ok';}else{history.go(0);}"/></td>
</tr>
</form>
<%shu=shu+cint(rs("cpsl"))
heji=heji+cint(cp_zkj*rs("cpsl"))
rs.movenext
wend
rs.close
set rs=nothing%>
<% set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [ogwc] where sfdd=0 and userid="&session("username")&" order by id desc",conn,1,1
if rs.eof and rs.bof then
response.Write""
else%>
<tr>
    <td colspan="3" align="center">合计</td>
    <td align="center"><%=shu%></td>
    <td align="center">&nbsp;</td>
    <td align="center"><%=heji%></td>
    <td colspan="2" align="center">&nbsp;</td>
  </tr>
<%end if %>


</table>
<% set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [ogwc] where sfdd=0 and userid="&session("username")&" order by id desc",conn,1,1
if rs.eof and rs.bof then
response.Write ""
else
while not rs.eof
cpidh=cpidh&rs("cpid")&","
rs.movenext
wend%>
<form name="alipayment" action="/alipay/listPayok.asp" method="get">
<table width="100%"  border="0" cellpadding="6" cellspacing="0" style="border-collapse:collapse;line-height:22px; text-indent:3px; margin-top:10px;background: #e5e5e5;">
  <tr>
    <td align="center">合计金额：<b style="color:#F00; font-size:20px">￥<%=heji%>.00</b> 元
    <input name="out_trade_no" type="hidden" size="5"  value="<%=BuyId%>"/>
    <input name="WIDtax" type="hidden" size="5"  value="0"/>
    <input name="WIDFreight" type="hidden" size="5"  value="0"/>
    <input name="cpid" type="hidden" size="5"  value="<%=left(cpidh,len(cpidh)-1)%>"/>
    <input name="userid" type="hidden" size="5"  value="<%=userid%>"/>
    <input name="rmb" type="hidden" size="5"  value="<%=heji%>"/>
    <input name="quantity" type="hidden" value="1" size="10"/>
    <input name="WIDshow_url"type="hidden"  value="<%=zych_url%>/Products/" size="30" readonly/></td>
    <td width="150" align="right"><a href="/Products/" id="btn1" class="btn2" style="margin:0px 5px;">继续选购</a></td>
    <td width="150" align="right"><a href="javascript:;" id="btn1" class="btn2" onClick="TopicShow(this,'dingdang')">支付结算</a></td>
  </tr>
</table>
<div id="dingdang" style="display:none">
<ul>
<li><p>单位名称</p><dfn><input name="title" type="txet" size="25"  value="<%=usergs%>" placeholder="请填入单位名称.." maxlength="20"/></dfn></li>
<li><p>联系电话</p><dfn><input name="WID_tel" type="txet" size="25"  value="<%=usertel%>" placeholder="请填入单位联系电话.." maxlength="11"/></dfn></li>
<li><p>联系人名</p><dfn><input name="WID_name" type="txet" size="25"  value="<%=username%>" placeholder="请填入联系人姓名.." maxlength="4"/></dfn></li>
<li><p>联系手机</p><dfn><input name="phone" type="txet" size="25"  value="<%=usersj%>" placeholder="请填入联系人联系电话.." maxlength="11"/></dfn></li>
<li><p>联系Q Q</p><dfn><input name="qq" type="txet" size="25"  value="<%=userqq%>" placeholder="请填入联系人联系电话.." maxlength="11"/></dfn></li>
<li><p>配送地址</p><dfn><input name="WIDreceive_address" type="txet" size="25"  value="<%=useradd%>" placeholder="请填入配送地址.." maxlength="30"/></dfn></li>
</ul>
<button type="submit" class="btn">提交订单</button>
</div>
</form>
<% end if
rs.close
set rs=nothing %>
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

<!--#include file="../Include/bottom.asp" -->
</body>
</html>
<%if Request.QueryString("add")="order" then
spid=request.form("spid")

'新加订单
set rs=server.createobject("adodb.recordset")
sql="select * from [Prodd]"
rs.open sql,conn,1,3
OrderNo=request.form("OrderNo")
title=request.form("title")
userid=request.form("userid")
tel=request.form("tel")
address=request.form("address")
name=request.form("name")
phone=request.form("phone")
sm=request.form("sm")
spid=Replace(Trim(request.form("spid"))," ","")
VerifyCode=request.form("VerifyCode")
if title=""  then 
response.Write("<script language=javascript>alert('公司名称不能为空!');history.go(-1)</script>") 
response.end 
end if
if tel=""  then 
response.Write("<script language=javascript>alert('请填写公司联系电话!');history.go(-1)</script>") 
response.end 
end if
if address=""  then 
response.Write("<script language=javascript>alert('请填写公司联系地址!');history.go(-1)</script>") 
response.end 
end if
if name=""  then 
response.Write("<script language=javascript>alert('请填写联系人姓名!');history.go(-1)</script>") 
response.end 
end if
if phone=""  then 
response.Write("<script language=javascript>alert('请填写联系人电话!');history.go(-1)</script>") 
response.end 
end if
rs.addnew
rs("OrderNo")=OrderNo
rs("title")=title
rs("spid")=spid
rs("userid")=userid
rs("tel")=tel
rs("address")=address
rs("name")=name
rs("phone")=phone
rs("rmb")=heji
rs("sm")=sm
rs.update
rs.close
set rs=nothing
'修改进货单状态
set rst=server.CreateObject("adodb.recordset")
rst.open "select * from [ProList] where userid="&userid&" and dd_id=0 order by id asc",conn,1,1
while not rst.eof
conn.execute("update ProList set dd_id=1 where userid="&userid&" and spid in ("&rst("spid")&")")
conn.execute("update Products set spkc=spkc-"&rst("spsl")&" where id in ("&rst("spid")&")")
rst.movenext
wend
rst.close
set rst=nothing
response.Write("<script language=""javascript"">alert(""订单成功！"");window.location.href='my_order.asp';</script>")
end if%>

<%if Request.QueryString("xiugai")="ok" then
id=request("id") 
set rs=server.CreateObject("adodb.recordset")
sql="select * from [ogwc] where id="&id 
rs.open sql,conn,1,3
cpsl=request.form("cpsl")  
IF not isNumeric(request("spsl")) then
response.write("<script>alert(""数量必须为数字！""); history.go(-1);</script>")
response.end
end if
rs("cpsl")=cpsl
rs.update 
rs.close 
response.Write("<script language=""javascript"">alert(""当前数量修改成功！"");window.location.href='my_addlist.asp';</script>")
end if
'删除
if request("del")="ok" then
id=Request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from [ogwc] where id="&id
rs.open sql,conn,2,3
rs.delete
rs.update
Response.Write "<script>alert('当前订单删除成功！');window.location.href='my_addlist.asp';</script>"
end if
%>