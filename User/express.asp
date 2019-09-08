<!--#include file="../Include/Inc.Asp" -->
<%Response.Charset="gb2312" 
Server.ScriptTimeout = 999999999

com = Request("com")'公司
nu = Request("nu")'单号

'Response.Write SendURL
'请勿删除版权提示信息，违者取消API功能，谢谢合作。
myxml="http://api.kuaidi.com/openapi.html?id="&kdappKey&"&com="&com&"&nu="&nu&"&show=1&muti=0&order=desc"
set xml = server.CreateObject("Microsoft.XMLDOM")
xml.async = "false"
xml.resolveExternals = "false"
xml.setProperty "ServerHTTPRequest", true
xml.load(myxml)
Set objNodes = xml.getElementsByTagName("data")
 Response.Write("<ul class=""timeline"">")
For i = 0 to objNodes.length -1 
if i=0 then na="time_active" else na="time_normal"
if i=0 then nb=" class=font" else nb=""
    Response.write "<li "&nb&"><div class="""&na&"""></div><div class=""kdbody""><span class=""kjiao""></span><span class=""time"">"&Trim(objNodes(i).selectSingleNode("time").Text)&"</span><span class=""title"">"&Trim(objNodes(i).selectSingleNode("context").Text)&"</span></div></li>"
Next 
Response.Write("</ul>")
'输入查询结果
Response.Write("<p>查询数据由：KuaiDi.Com 网站提供</p>")
%>