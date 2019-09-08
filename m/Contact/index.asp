<!--#include file="../../Include/conn.asp"-->
<!--#include file="../Include/config.asp" -->
<!--#include file="../Include/Sql.Asp" -->
<!--#include file="../Include/zych_sql.asp"-->
<!DOCTYPE HTML>
<html lang="en">
<head>
<title><%=zych_class(murl)%>_<%=zych_home%></title>
<meta charset="gb2312">
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<link rel="stylesheet" type="text/css" href="../css/style.css" media="screen" />
<script type="text/javascript">
var messageDelay = 2000; 
$( init );
function init() {
  $('#contactForm').submit( submitForm ).addClass( 'positioned' );
  $('a[href="#contactForm"]').click( function() {
    $('#content').fadeTo( 'slow', .2 );
    $('#contactForm').fadeIn( 'slow', function() {
      $('#senderName').focus();
    } )

    return false;
  } );
  // When the "Escape" key is pressed, close the form
  $('#contactForm').keydown( function( event ) {
    if ( event.which == 27 ) {
      $('#contactForm').fadeOut();
      $('#content').fadeTo( 'slow', 1 );
    }
  } );

}
// Submit the form via Ajax
function submitForm() {
  var contactForm = $(this);
  // Are all the fields filled in?
  if ( !$('#senderName').val() || !$('#sendertitle').val() || !$('#senderEmail').val() || !$('#message').val() ) {
    // No; display a warning message and return to the form
    $('#incompleteMessage').fadeIn().delay(messageDelay).fadeOut();
    contactForm.fadeOut().delay(messageDelay).fadeIn();
  } else {
    // Yes; submit the form to the PHP script via Ajax
    $('#sendingMessage').fadeIn();
    $.ajax( {
      url: contactForm.attr( 'action' ) + "@ajax=true",
      type: contactForm.attr( 'method' ),
      data: contactForm.serialize(),
      success: submitFinished
    } );
  }
  // 防止默认表单提交的发生
  return false;
}
//处理Ajax响应
function submitFinished( response ) {
  response = $.trim( response );
  $('#sendingMessage').fadeOut();
  if ( response == "success" ) {
    $('#successMessage').fadeIn().delay(messageDelay).fadeOut();
    $('#sendertitle').val( "" );
	$('#senderName').val( "" );
    $('#senderEmail').val( "" );
    $('#message').val( "" );
	$('#VerifyCode').val( "" );
    $('#content').delay(messageDelay+500).fadeTo( 'slow', 1 );
  } else {
    $('#failureMessage').fadeIn().delay(messageDelay).fadeOut();
    $('#contactForm').delay(messageDelay+500).fadeIn();
  }
}

</script>
</head>
<body>
<div id="wrapper">
<!--#include file="../Include/top.asp"--> 
<section id="main">
<!--Block Starts-->
<div class="block_module paper_bh_white">
  <h2>联系我们</h2>
  <ul id="contact">
    <li class="address"><%=zych_dz%></li>
    <li class="telephone"><strong><%=zych_tel%></strong></li>
    <li class="website"><a href="<%=zych_url%>"><%=zych_url%></a></li>
  </ul>
</div>
<!--Block Ends-->
<div class="block_module paper_bh_white">
<h2>访客留言</h2>
<ul id="comments">
<% set rs=server.createobject("adodb.recordset") 
exec="select top 5 * from ly where sh=1 order by id desc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div clacc=""onxx"">暂无留言记录！</div>"
end if
while not rs.eof
if IsNull(rs("hf")) or trim(rs("hf")&"")="" then
adminhf="<font color=#FF0000>此留言暂未回复！</font>"
else
adminhf="<font color=#084B8A>"&rs("hf")&"</font>"
end if%>
<li class="odd">
    <h5><a href="#"><%=rs("name")%><img alt="" src="../images/user_admin.jpg"></a></h5>
    <div class="comment_body"> <%=stvalue(DelHtml(rs("body")),200)%><span><%=rs("data")%></span> </div>
</li>
<li class="even">
    <div class="comment_body">管理员回复: <%= adminhf %> </div>
</li>
<%rs.movenext
wend
rs.close
set rs=nothing%>

</ul>
</div>
<!--Comments Starts-->
<div class="block_module comment_form">
  <h2>在线留言</h2>
  <form id="contactForm" action="?action=bookadd" method="post">
    <ul id="moby_form">
      <li><input type="text" name="sendertitle" id="sendertitle" placeholder="留言标题" required/></li>
      <li><input type="text" name="senderName" id="senderName" placeholder="您的昵称" required/></li>
      <li><input type="email" name="senderEmail" id="senderEmail" placeholder="您的邮箱" required /></li>
      <li><textarea name="message" id="message" placeholder="您想要说的内容" required cols="80" rows="10" maxlength="500"></textarea></li>
      <input name="sh" type="hidden" value="<%=config("booksh")%>" size="10"/>
      <li style="height:40px;"><input type="text" name="VerifyCode" id="VerifyCode" style="width:100px;float:left" placeholder="验证码" required /> <img src="../Include/safecode.asp" alt="看不清,请点击刷新验证码!" class="v" width="70" height="28" onClick="this.src='../Include/safecode.asp?t='+(new Date().getTime());"><input type="submit" id="sendMessage" name="sendMessage" style="float:right" value="确认提交" /></li>
    </ul>
  </form>      
  <div id="sendingMessage" class="statusMessage">
    <p>发送你的信息。请您稍候。</p>
  </div>
  <div id="successMessage" class="statusMessage">
    <p>发送消息的感谢！我们会尽快给您。</p>
  </div>
  <div id="failureMessage" class="statusMessage">
    <p>有一个问题，发送你的信息。请再试一次。</p>
  </div>
  <div id="incompleteMessage" class="statusMessage">
    <p>请填写所有领域的形式发送前。</p>
  </div>
</div>
<!--Comments Ends-->
</section>
<!--Section Ends-->
<!--Footer Starts-->
<!--#include file="../Include/bottom.asp" -->
<!--Footer Ends-->
</div>
</body>
</html>
<%
if request("action")="bookadd" then
set rs=server.createobject("adodb.recordset")
sql="select * from ly"
rs.open sql,conn,1,3
title=request.form("sendertitle")
body=request.form("message")
name=request.form("senderName")
qq=request.form("qq")
mail=request.form("senderEmail")
tel=request.form("tel")
sh=request.form("sh")
VerifyCode=request.form("VerifyCode") 
IF title="" then
response.write("<script>alert(""标题不能为空""); history.go(-1);</script>")
response.end
end if
IF name="" then
response.write("<script>alert(""昵称不能为空""); history.go(-1);</script>")
response.end
end if
IF mail="" then
response.write("<script>alert(""邮箱不能为空""); history.go(-1);</script>")
response.end
end if
IF body="" then
response.write("<script>alert(""留言内容不能为空""); history.go(-1);</script>")
response.end
end if
IF VerifyCode="" then
response.write("<script>alert(""验证码不能为空""); history.go(-1);</script>")
response.end
end if
if cstr(Session("firstecode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('验证码错误!');history.go(-1)</script>")
response.End
end if
IF sh="" or not isNumeric(request("sh")) then
response.write("<script>alert(""警告！请勿尝试外部提交数据！""); history.go(-1);</script>")
response.end
end if
rs.addnew
rs("title")=title
rs("body")=body
rs("name")=name
rs("qq")=qq
rs("mail")=mail
rs("tel")=tel
rs("sh")=sh
rs.update
rs.close
set rs=nothing
conn.close
set rs=nothing
if sh<>1 then
Response.Write "<script>alert('留言提交需要管理员审核才能显示！');window.location.href='?';</script>" 
else 
Response.Write "<script>alert('留言发布成功！');window.location.href='?';</script>" 
end if
end if
%>