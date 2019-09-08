<%if body_no<>0 then
response.write("<body style=""background:url("&body_bj&") repeat-x top "&body_color&""">")
end if%>
<% If wap=0 Then %>
<script type ="text/javascript">var _hmt = _hmt || [];(function() {	var ua = navigator.userAgent;	if (/iphone|ipad|android/i.test(ua)) {	window.location.href = "<%=Dir%>m";	}})();</script>
<% End If %>
<div class="top" style="display:none;">
<div class="box"><div class="time">今天是:<span id="webjx"></span></div>
<script>setInterval("webjx.innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt(new Date().getDay());",1000);</script>
<div class="tnav"><%=zych_Navigation_top%> <%=zych_login%></div></div>
</div>
<div class="head">
<div class="top_bar">
<div class="logo"><img src="<%=zych_logo%>" width="<%=zych_logoiwidth%>" height="<%=zych_logoheight%>"/></div>
<div class="bogorht"><div class="at1"><%=zych_tel%></div><div class="at2"><%=zych_logoright%></div></div>
<div class="clear"></div>
</div>
</div></div>
<!-- 导航代码end -->
<div id="navigation"><div class="nav">
<ul class="nav_list"><%=zych_Navigation %></ul>
</div></div>
<!-- 导航代码begin -->