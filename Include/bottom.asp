<div id="bottom">
<div class="copyright">
  <div class="text clear">
  <div class="info">
    <p><%=ditu_mc%></p>
	<p><%=zych_dz%></p>
	<p><%=zych_beian%></p>
    <p><%=zych_copyright%></p>  
  </div>

<div class="icon">
  <a><span class="sitemap"></span><i>网站地图</i></a>
  <a><span class="wechat"></span><i>微信网站</i></a>
  <a><span class="mail"></span><i>企业邮箱</i></a>
  </div>
  
    <div class="wx">
  <img src="images/wx.jpg" />
  </div>
  </div>  
  

  
  
  
</div>
<div class="clear"></div>
</div>
<%if qqkf=0 then %>
<div id="qq_zych" onmouseover=toBig() onmouseout=toSmall()>
<div id="qq_service">
<div class="head"><p>在线客服</p></div>
<div class="info">
<div class="box yh">
<ul><%=zych_qq%><li><a href="<%=dir%>Contact/">在线留言</a></li></ul>
<% If zych_rwmon<>1 Then %>
<div class="rwm"><img src="<%=zych_rwmurl%>" alt="<%=zych_home%>" width="120" height="120" title="用手机微信扫一扫" /></div>
<% End If %>
</div>
</div>
</div>
<div class="Obtn"></div>
</div>
<SCRIPT language=javascript src="<%=dir%>js/pfkf1.js" type=text/javascript></SCRIPT>
<% End If %>
