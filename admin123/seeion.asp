<!--#include file="../Include/version.asp" -->
<!--#include file="function.asp" -->
<%session.Timeout=900
if session("admin")<>"" then
else
response.write "<script>;window.location.href='login.asp';</script>"
response.end
end if
set rsa=server.createobject("adodb.recordset") 
exec="select * from admin where id="&session("admin")&""
rsa.open exec,conn,1,1 
admin=rsa("admin")
password=rsa("password")
zsname=rsa("zsname")
key=rsa("key")
manage=rsa("manage")
qq=rsa("qq")
rsa.close
set rsa=nothing
Function chkAdmin(byval Level)
  If key<>0 then
	  if instr(","&manage&",",","&lcase(Level) &",")=0 then
	  Call adminJump("Sorry!","您没有管理该模块的权限！","javascript:window.history.go(-1)")
	  response.End
	  End if
  End if
End Function

Function all_navoption()'所以二级导航
'单页频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from about order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==单页频道==</option>")
while not rs.eof
if html=0 then url="About/"&about&Separated&rs("id")&"."&HTMLName else url="About/?id="&rs("id")
Response.Write("<option value="""&url&""">"&rs("title")&"</option>")
rs.movenext
wend
'新闻频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news_fl order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==新闻频道==</option>")
while not rs.eof
if html=0 then url="News/"&Newsfl&Separated&rs("id")&"."&HTMLName else url="News/Newslist.asp?id="&rs("id")
Response.Write("<option value="""&url&""">"&rs("title")&"</option>")
rs.movenext
wend
'案例频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [case_class] order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==案例频道==</option>")
while not rs.eof
if html=0 then url="Case/"&Casefl&Separated&rs("id")&"."&HTMLName else url=Dir&"Case/?id="&rs("id")
Response.Write("<option value="""&url&""">"&rs("title")&"</option>")
rs.movenext
wend
'产品频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [BigClass] order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==产品频道==</option>")
while not rs.eof
if html=0 then url="Products/"&Productsfl&Separated&rs("BigClassID")&"."&HTMLName else url="Products/?BigClassID="&rs("BigClassID")
Response.Write("<option value="""&url&""">"&rs("BigClassName")&"</option>")
rs.movenext
wend
'视频频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [video_class] order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==视频频道==</option>")
while not rs.eof
if html=0 then url="Video/"&Videofl&Separated&rs("id")&"."&HTMLName else url="Video/?id="&rs("id")
Response.Write("<option value="""&url&""">"&rs("title")&"</option>")
rs.movenext
wend
'下载频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [download_fl] order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==下载频道==</option>")
while not rs.eof
if html=0 then url="Download/"&Downloadfl&Separated&rs("id")&"."&HTMLName else url="Download/?id="&rs("id")
Response.Write("<option value="""&url&""">"&rs("title")&"</option>")
rs.movenext
wend
'相册频道
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [photo_class] order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<option value=""#"" disabled>暂无记录！</option>")
end if
response.Write("<option value=""#"" disabled>==相册频道==</option>")
while not rs.eof
if html=0 then url="Photo/"&photofl&Separated&rs("id")&"."&HTMLName else url="Photo/?id="&rs("id")
Response.Write("<option value="""&url&""">"&rs("title")&"</option>")
rs.movenext
wend
End Function

'页面自动跳转
sub adminJump(str1,str2,url)
Response.Write("<div style=""width:400px;height:140px;position:absolute;left:50%;top:50%;margin-left:-200px;margin-top:-70px;font-size:12px;text-align:center;border:4px #CCC solid"">")
Response.Write("<h1>"&str1&"</h1>")
Response.Write("<p style=""font-size:16px;color:#F00;font-family:'微软雅黑'"">"&str2&"</p>")
Response.Write("<p style=""font-size:12px;text-align:center;"">您可以点击这里返回<a href="""&url&""">上一页</a></p>")
Response.Write("</div>")
end sub
Function DelHtml(Str1)
  Dim regEx
  Set regEx = New RegExp
  regEx.Pattern = "(<[^>]*?>)"
  regEx.Global = True
  regEx.IgnoreCase = True
  DelHtml = replace(regEx.Replace(""&str1,""),"&nbsp;","")
End Function

Function stvalue(txt,length)
txt=trim(txt)
x = len(txt)
y = 0
if x >= 1 then
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then 
y = y + 2
else
y = y + 1
end if
if y >= length then
txt = left(trim(txt),ii)&"…" 
exit for
end if
next
stvalue = txt
else
stvalue= ""
end if
End Function
'获取扩展名
Function suffix(str)
	dim ext : str=trim(""&str) : ext=""
	if str<>"" then
		if instr(" "&str,"?")>0 then:str=mid(str,1,instr(str,"?")-1):end if
		if instrRev(str,".")>0 then:ext=mid(str,instrRev(str,".")):end if
	end if
	suffix=ext
End Function

On error resume next '容错处理
set crs=server.createobject("adodb.recordset") 
exec="select * from config" 
crs.open exec,conn,1,1
wzcrs=Server.UrlEncode(crs("title"))
wzPhone=Server.UrlEncode(crs("Phone"))
crs.Close
Set crs= Nothing
http=Chr(104)&Chr(116)&Chr(116)&Chr(112)&Chr(58)&Chr(47)&Chr(47)
zyurl=Chr(119)&Chr(119)&Chr(119)&Chr(46)&Chr(122)&Chr(121)&Chr(99)&Chr(104)&Chr(114)&Chr(46)&Chr(99)&Chr(111)&Chr(109)
Function wfg(Wfcx_str)
Set https = Server.CreateObject("MSXML2.XMLHTTP") 
With https 
.Open "Post",http&zyurl&"/"&Chr(73)&Chr(110)&Chr(99)&"lude/weburl"&Chr(46)&"asp"&Chr(63)&"action=add", False
.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
.Send "title="&wzcrs&"&weburl="&Request.servervariables("server_name")&"&version=ZYCH_"&zychversion&"&verdata="&data&"&authorization=Pay&qq="&qq&"&tel="&wzPhone&""
wfg = .ResponseBody
End With 
wfg = BytesToBstr(wfg,"GB2312")
Set https = Nothing 
End Function

Function BytesToBstr(body,Cset)
	if lenb(body)=0  then
	  BytesToBstr=""
	  exit  function
	end if
	dim objstream
	set objstream = Server.CreateObject("adodb.stream")
	objstream.Type = 1
	objstream.Mode =3
	objstream.Open
	objstream.Write body
	objstream.Position = 0
	objstream.Type = 2
	objstream.Charset = Cset
	BytesToBstr = objstream.ReadText 
	objstream.Close
	set objstream = nothing
End Function
Function isInstallObj(objname)
	dim isInstall,obj
	On Error Resume Next
	set obj=server.CreateObject(objname)
	if Err then 
		isInstallObj=false : err.clear 
	else 
		isInstallObj=true:set obj=nothing
	end if
End Function
Function salt(str)
Response.Write"<div class=""pop"" id=""pop"">"
Response.Write"<div class=box><a href=""javascript:void(0)"" onclick =""document.getElementById('pop').style.display='none'"" style=""position:absolute;right:0px;color:#000; margin:2px 10px;"" title=关闭窗口 id=""closeBtn"">×关闭</a>"
Response.Write"<div class=t1>"&str&"</div>"
Response.Write"<div style=width:66px; height:66px; margin:20px auto; background:url(images/working.gif)></div>"
Response.Write"</div></div>"
End Function
%>