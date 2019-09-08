<!--#include file="Include/conn.asp"-->
<!--#include file="Include/Sql.Asp" -->
<%set rsc=server.createobject("adodb.recordset") 
exec="select * from config" 
rsc.open exec,conn,1,1
zych_home=""&rsc("title")&""
zych_keywords=""&rsc("keywords")&""
zych_description=""&rsc("description")&""
zych_off=rsc("off_sm")
rsc.close
set rsc=nothing%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Language" content="zh-cn">
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<title><%=zych_home%></title>
<meta name="keywords" content="<%=zych_keywords%>" />
<meta name="description" content="<%=zych_description%>" />
<style type="text/css">
body{ font-size:12px; background:#F2F2F2}
#mian{width:700px;height:400px;position:absolute;background:#FFF;left:50%;top:50%;margin-left:-350px;margin-top:-200px; border:1px #999 dashed;}
.ts_404{width:600px;height:220px; margin:0px auto;font-size:160px; text-align:center; color:#999;font-family:'Microsoft Sans Serif' }
.ts_1{height:30px; margin:0px auto;font-size:30px; text-align:center; color:#09F;font-family:'微软雅黑'}
.ts_2{height:100px; margin:0px auto;font-size:14px; line-height:100px; text-align:center; color:#09F;font-family:'微软雅黑'}
.Copyright{font-size:12px; line-height:50px; text-align:center; color:#09F; font-family:'微软雅黑'}
.Copyright a{color:#09F;text-decoration: none;}
</style>
</head>
<body>
<div id="mian">
<div class="ts_404">Sorry!</div>
<div class="ts_1"><%=zych_off%></div>
<div class="ts_2">给您带来的不便敬请谅解，我们将尽快恢复网站！</div>
<div class="Copyright">Copyright &copy; Powered by <a href="http://www.zychr.com" target="_blank">自由策划</a>. All Right Reserved. </div>
</div>
</body>
</html>