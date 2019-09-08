<!--#include file="../../Include/Inc.Asp" -->
<%
set config=server.createobject("adodb.recordset") 
exec="select * from config" 
config.open exec,conn,1,1 
if config("on") = 1 then
response.write (""&config("off_sm")&"")
response.End
else 
end if 
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
	stvalue = ""
end if
End Function
'获取扩展名
Function getFileFormat(str)
	dim ext : str=trim(""&str) : ext=""
	if str<>"" then
		if instr(" "&str,"?")>0 then:str=mid(str,1,instr(str,"?")-1):end if
		if instrRev(str,".")>0 then:ext=mid(str,instrRev(str,".")):end if
	end if
	getFileFormat=ext
End Function
function weburl()'网站URL
If Request.ServerVariables("SERVER_PORT") = "80" Then
GetSiteUrl = "http://" & Request.ServerVariables("server_name")
Else
GetSiteUrl = "http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("SERVER_PORT")
End If
end function
wz=request.ServerVariables("HTTP_HOST")&request.ServerVariables("URL")
wurl=right(left(wz,instrrev(wz,"/")),(len(left(wz,instrrev(wz,"/")))-len(request.ServerVariables("HTTP_HOST")))-1)
murl=right(wurl,len(wurl)-2)
function zych_class(murl)'菜单导航调用参数
set nav=server.CreateObject("adodb.recordset")
nav.open "select * from menu where url='"&murl&"'",conn,1,1
if nav.eof and nav.bof then
response.Write ""
else
response.Write nav("title")&""
end if
nav.close
set nav=nothing
end function

function zych_en_class(wurl)'菜单导航调用参数
set nav=server.CreateObject("adodb.recordset")
nav.open "select * from menu where url='"&wurl&"' ",conn,1,1
if nav.eof and nav.bof then
response.Write ""
else
response.Write nav("en_title")&""
end if
nav.close
set nav=nothing
end function

function zych_login()'会员登陆调用参数
if session("username")<>"" then 
set rss=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rss.open exec,conn,1,1 
response.write("<font><a href="""&Dir&"member/index.asp"">["&rss("useradmin")&"]</a>&nbsp;欢迎您!<a href="""&Dir&"member/LoginOut.asp""> 退出</a>&nbsp;</font>")
else
response.write("<font color=""#FFFFFF""><a href="""&Dir&"member/login.asp"">[登陆]</a>&nbsp;</font>")
end if
end function

function zych_home()
zych_home=""&config("title")&zych_c&""
end function
function zych_keywords()
zych_keywords=""&config("keywords")&""
end function
function zych_description()
zych_description=""&config("description")&""
end function
function zych_url()
zych_url=""&config("url")&""
end function
function zych_logo()
zych_logo=""&config("logo")&""
end function
function zych_logoiwidth()
zych_logoiwidth=""&config("logoiwidth")&""
end function
function zych_logoheight()
zych_logoheight=""&config("logoheight")&""
end function
function zych_name()
zych_name=""&config("name")&""
end function
function zych_css()
zych_css=""&config("css")&""
end function
function zych_mail()
zych_mail=""&config("mail")&""
end function
function zych_Phone()
zych_Phone=""&config("Phone")&""
end function
function zych_tel()
zych_tel=""&config("tel")&""
end function
function zych_fax()
zych_fax=""&config("fax")&""
end function
function zych_Postal()
zych_Postal=""&config("Postal")&""
end function
function zych_dz()
zych_dz=""&config("dz")&""
end function
function zych_copyright()
zych_copyright=""&config("copyright")&""
end function
function zych_logoright()
zych_logoright=""&config("logoright")&""
end function
function zych_daohangnr()
zych_daohangnr=""&config("daohangnr")&""
end function
function zych_beian()
zych_beian="<a href=""http://www.miibeian.gov.cn""  target=""_blank"">"&config("beian")&"</a>"
end function
function zych_video_sz()
zych_video_sz=""&config("video_sz")&""
end function
function zych_video_wz()
zych_video_wz=""&config("video_wz")&""
end function
function zych_zztj()
zych_zztj=""&config("zztj")&""
end function
function zych_foot_left()
zych_foot_left=""&config("foot_left")&""
end function
function zych_Visits()
sql="update config set js=js+1 where js is not null" 
conn.execute(sql) 
response.Write(""&config("bq")&" 访问量："&config("js")&"")
end function
call zych_Nav()
zych_c=""&Chr("32")&Chr("45")&Chr("80")&Chr("111")&Chr("119")&Chr("101")&Chr("114")&Chr("101")&Chr("100")&Chr("32")&Chr("98")&Chr("121")&Chr("32")&Chr("122")&Chr("121")&Chr("99")&Chr("104")&Chr("114")&Chr("46")&Chr("99")&Chr("111")&Chr("109")&""
function zych_Nvachannel(ID)
dim zych_channel
set rss=server.CreateObject("adodb.recordset")
sql="select * from menu where id="&ID&""
rss.open sql,conn,1,1
if rss.eof and rss.bof then
response.Write("未知频道")
else
Response.Write(rss("title"))
end if
end function%>

<%
function zych_about_list()'单页列表部分调用参数
set aboutlist=server.CreateObject("adodb.recordset")
aboutlist.open "select * from about where xianshi=1 order by px_id asc",conn,1,1
if aboutlist.eof and aboutlist.bof then
response.Write("<a>暂无记录</a>")
end if
while not aboutlist.eof
if html=0 then
response.Write("<li><a href="""&Dir&"About/"&about&""&Separated&""&aboutlist("id")&"."&HTMLName&""" title="""&aboutlist("title")&""">"  & aboutlist("title") & "</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"About/?id="&aboutlist("id")&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a></li>")
end if
aboutlist.movenext
wend
aboutlist.close
set aboutlist=nothing
end function


function zych_about_foot()'底部单页列表部分调用参数
set aboutlist=server.CreateObject("adodb.recordset")
aboutlist.open "select * from about where xianshi=1 order by px_id asc",conn,1,1
if aboutlist.eof and aboutlist.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not aboutlist.eof

response.Write("<a href="""&Dir&"m/About/Show.asp?id="&aboutlist("id")&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a> | ")
aboutlist.movenext
wend
aboutlist.close
set aboutlist=nothing
end function

function zych_about_list_pd()'单页列表部分调用参数
set aboutlist=server.CreateObject("adodb.recordset")
aboutlist.open "select * from about where xianshi=1 order by px_id asc",conn,1,1
if aboutlist.eof and aboutlist.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not aboutlist.eof
if html=0 then
response.Write("<li><a href="""&Dir&"About/"&about&""&Separated&""&aboutlist("id")&"."&HTMLName&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"About/?id="&aboutlist("id")&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a></li>")
end if
aboutlist.movenext
wend
aboutlist.close
set aboutlist=nothing
end function


function zych_News_fl_list()'新闻分类调用参数
set rsfl=server.CreateObject("adodb.recordset")
rsfl.open "select * from news_fl order by px_id asc",conn,1,1
if rsfl.eof and rsfl.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not rsfl.eof
if html=0 then
response.Write("<li><a href="""&Dir&"News/"&Newsfl&""&Separated&""&rsfl("id")&"."&HTMLName&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"News/Newslist.asp?id="&rsfl("id")&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
end if
rsfl.movenext
wend
rsfl.close
set rsfl=nothing
end function

function zych_Navigation()'菜单导航调用参数
dim Navigation
set Navigation=server.CreateObject("adodb.recordset")
Navigation.open "select * from menu where yc=1 order by px_id asc",conn,1,1
while not Navigation.eof
if Navigation("url")=wurl then
no=" class=""on"""
else
no=""
end if
Response.Write("<i></i><li><a"&no&" href="""&dir&""&Navigation("url")&""" rel='"&Navigation("id")&"' target="""&Navigation("openfs")&""">"&Navigation("title")&"</a></li>")
Navigation.movenext
wend
Response.Write("<i></i>")
Navigation.close
set Navigation=nothing
end function

function zych_Navigation_1()'菜单导航调用参数

set na=server.CreateObject("adodb.recordset")
na.open "select * from menu order by px_id asc",conn,1,1
while not na.eof

set nb=server.CreateObject("adodb.recordset")
nb.open "select * from menu_nav where ssfl="&na("id")&" order by px_id asc",conn,1,1
if nb.eof and nb.bof then
else
Response.Write("<ul id="""&na("id")&""" class=""dropMenu"">")
while not nb.eof
Response.Write("<li><a href="""&nb("url")&""" title="""&nb("title")&""">"&nb("title")&"</a></li>")
nb.movenext
wend
Response.Write("</ul>")
end if
na.movenext
wend
na.close
set na=nothing
end function
%>


<%
function zych_Navigation_top()'网站顶部菜单导航调用参数
dim Navigation
set Navigation=server.CreateObject("adodb.recordset")
Navigation.open "select * from menu where yc=2 order by px_id asc",conn,1,1
while not Navigation.eof
response.Write("<a href="""&dir&""&Navigation("url")&""" target="""&Navigation("openfs")&""">"&Navigation("title")&"</a>"&vbcrlf)
Navigation.movenext
wend
Navigation.close
set Navigation=nothing
end function
%>
<%function zych_Flash()'Flash幻灯调用参数%>
<SCRIPT language=JavaScript type=text/javascript>
var swf_width='980';
var swf_height='200';
var configtg='0xffffff:文字颜色|2:文字位置|0x000000:文字背景颜色|10:文字背景透明度|0xffffff:按键文字颜色|0x4f6898:按键默认颜色|0x000033:按键当前颜色|6:自动播放时间|3:图片过渡效果|1:是否显示按钮|_blank:打开新窗口';
<%
		set db=conn.execute("select * from [flash] order by px_id asc")
			i=0
			do while not db.eof
				 files=files&"|"&db("img")
				 links=links&"|"&db("link")
				 texts=texts&"|"&db("title")
				db.moveNext
				i=i+1
			loop
if db.eof and db.bof then
response.write "var files=''"&vbcrlf
response.write "var links=''"&vbcrlf
response.write "var texts=''"&vbcrlf
else
response.write "var files='"&right(files,len(files)-1)&"'"&vbcrlf
response.write "var links='"&right(links,len(links)-1)&"'"&vbcrlf
response.write "var texts='"&right(texts,len(texts)-1)&"'"&vbcrlf
end if
%>
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="<%=dir%>images/slideflash.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'">');
document.write('<embed src="<%=dir%>images/slideflash.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%end function%>



<%function zych_index_video()'首页视频调用参数%>
<%if flvswf=1 then
response.Write "<div class=""flv"">"%>
<script type="text/javascript">
var swf_width=280
var swf_height=170
<%
set db=conn.execute("select * from [video] where xianshi=1 order by px_id asc")
			i=0
			do while not db.eof
				 texts=texts&"|"&db("title")
				 files=files&"|"&db("body")
				db.moveNext
				i=i+1
			loop
if db.eof and db.bof then
response.write "texts='高屋置业'"&vbcrlf
response.Write "var files='http://www.shgwzy.com/ad/zych.flv'"&vbcrlf
else
response.write "texts='"&right(texts,len(texts)-1)&"'"&vbcrlf
response.write "var files='"&right(files,len(files)-1)&"'"&vbcrlf
end if
%>
var config='<%=zych_video_sz%>:自动播放|1:连续播放|100:默认音量|0:控制栏位置|2:控制栏显示|0x000033:主体颜色|60:主体透明度|0x66ff00:光晕颜色|0xffffff:图标颜色|0xffffff:文字颜色|<%=zych_video_wz%>:标题:logo地址|:结束swf地址'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="<%=weburl%><%=dir%>images/vcastr2.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'">');
document.write('<embed src="<%=weburl%><%=dir%>images/vcastr2.swf" wmode="opaque" FlashVars="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'" menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%response.Write "</div>"
end if
if flvswf=2 then
set rs=server.createobject("adodb.recordset") 
exec="select top 1 *  from [video] where xianshi=1 order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">无此视频！</div>"
response.End()
end if
lname=""&rs("body")&""
la=split(lname,"/")
num=ubound(la)
lb=split(la(num),".")
num2=ubound(lb)
suffix="."&lb(num2)&""
if suffix=".flv" then
response.Write "<div class=""flv"">"
%>
<script type="text/javascript">
var swf_width=280
var swf_height=170
texts='<%=rs("title")%>'
var files='<%=rs("body")%>'
var config='<%=zych_video_sz%>:自动播放|1:连续播放|100:默认音量|0:控制栏位置|2:控制栏显示|0x000033:主体颜色|60:主体透明度|0x66ff00:光晕颜色|0xffffff:图标颜色|0xffffff:文字颜色|<%=zych_video_wz%>:标题:logo地址|:结束swf地址'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="<%=weburl%><%=dir%>images/vcastr2.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'">');
document.write('<embed src="<%=weburl%><%=dir%>images/vcastr2.swf" wmode="opaque" FlashVars="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'" menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%else
response.Write "<div class=""swf"">"
response.Write"<embed src="""&rs("body")&""" allowFullScreen=""true"" quality=""high"" width=""280"" height=""170"" align=""middle"" allowScriptAccess=""always"" type=""application/x-shockwave-flash""></embed>"
end if
rs.close
set rs=nothing
response.Write "</div>"
end if
end function%>

<%
Function GetUrl(action)'加入收藏获取Url地址调用参数
GetUrl=request.servervariables("script_name")'赋目录及文件名
if action="div" then exit Function
GetUrl=Mid(Request.ServerVariables("script_name"),InstrRev(Replace(Request.ServerVariables("script_name"),"\","/"),"/")+1)'赋文件名
if action="page" then exit Function
GetUrl=request.servervariables("QUERY_STRING")
if action="action" then exit Function
GetUrl="http://"
GetUrl=GetUrl&request.servervariables("HTTP_HOST")'
if action="http" then exit Function
GetUrl=GetUrl&request.servervariables("script_name")'
if action="alldiv" then exit Function
if request.servervariables("QUERY_STRING")<>"" then GetUrl=GetUrl&"?"&request.servervariables("QUERY_STRING")'
End Function
%>


<%
function zych_download_fl()'下载分类调用参数
set downloadclas=server.CreateObject("adodb.recordset")
downloadclas.open "select * from download_fl order by px_id asc",conn,1,1
if downloadclas.eof and downloadclas.bof then
response.Write("&nbsp;暂无记录 !")
end if
while not downloadclas.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Download/"&Downloadfl&""&Separated&"" & downloadclas("id") & "."&HTMLName&""" title=""" & downloadclas("title") &""">"&downloadclas("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Download/?id=" & downloadclas("id")&""" title=""" &downloadclas("title")&""">"& downloadclas("title")&"</a></li>")
end if
downloadclas.movenext
wend
downloadclas.close
set downloadclas=nothing
end function

function zych_case_class()'手机案例分类调用参数
set ra=server.CreateObject("adodb.recordset")
ra.open "select * from case_class order by px_id asc",conn,1,1
if ra.eof and ra.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not ra.eof
response.Write("<li><a href=""?id=" & ra("id")&""" title=""" &ra("title")&""">"& ra("title")&"</a></li>")
ra.movenext
wend
ra.close
set ra=nothing
end function

function zych_Products_class()'手机产品分类
 set rs1=server.CreateObject("adodb.recordset")
rs1.open "select * from [bigClass] order by px_id asc",conn,1,1
if rs1.eof and rs1.bof then
response.Write("<li><a>暂无记录</a></li> ")
end if
while not rs1.eof
url="?BigClassID="&rs1("BigClassID")&""
response.write("<li><a href="""&url&""" title="&rs1("BigClassName")&">"&rs1("BigClassName")&"</a></li>"&vbcrlf)
set rs2=server.createobject("adodb.recordset") 
exec="select * from [SmallClass] where BigClassID="&rs1("BigClassID")&" order by px_id asc  " 
rs2.open exec,conn,1,1
if rs2.eof and rs2.bof then
response.Write("")
end if
do while not rs2.eof
url2="?BigClassID="&rs2("BigClassID")&"&amp;SmallClassID="&rs2("SmallClassID")&""
response.write("<li><a href="""&url2&""" title="&rs1("BigClassName")&">├─"&rs2("SmallClassName")&"</a></li>")
rs2.movenext
loop
rs2.close
set rs2=nothing
rs1.movenext
wend
rs1.close
set rs1=nothing
end function



function zych_photo_class()'相册分类调用参数
set photo=server.CreateObject("adodb.recordset")
photo.open "select * from photo_class order by px_id asc",conn,1,1
if photo.eof and photo.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not photo.eof
if html=0 then
response.Write("<a href="""&Dir&"Photo/"&photofl&""&Separated&"" & photo("id") & "."&HTMLName&""" title=""" & photo("title") &""">"  & photo("title") & "</a>")
elseif html=1 then
response.Write("<a href="""&Dir&"Photo/?id=" & photo("id") & """ title=""" & photo("title") &""">"  & photo("title") & "</a>")
end if
photo.movenext
wend
photo.close
set photo=nothing
end function


function zych_video_class()'视频分类调用参数
set video=server.CreateObject("adodb.recordset")
video.open "select * from video_class order by px_id asc",conn,1,1
if video.eof and video.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not video.eof
if html=0 then
response.Write("<li><a href="""&Videofl&""&Separated&"" & video("id") & "."&HTMLName&""" title=""" & video("title") &""">"  & video("title") & "</a><li>")
elseif html=1 then
response.Write("<li><a href=""?id=" & video("id") & """ title=""" & video("title") &""">"  & video("title") & "</a><li>")
end if
video.movenext
wend
video.close
set video=nothing
end function

function zych_index_case(Article)'首页推荐滚动图文调用参数
if Scroll=1 then
set tuijian=server.CreateObject("adodb.recordset")
tuijian.open "select top "&Article&" * from [case] where tuijian=1 order by id desc",conn,1,1
if tuijian.eof and tuijian.bof then
response.Write("<div class=""onxx"">暂无推荐案例记录</div>")
end if
while not tuijian.eof
if html=0 then
url=""&Dir&"Case/"&showcase&""&Separated&""&tuijian("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Case/ShowCase.asp?id="&tuijian("id")&""
end if
response.Write("<div class=""tui_box""><a href="""&url&""" title="""&tuijian("title")&"""><img src="""&tuijian("img")&""" width=""213"" height=""130"" alt="""&tuijian("title")&""" onerror=""this.src='images/nologo.jpg'""/><p>"&stvalue(tuijian("title"),24)&"</p></a></div>")
tuijian.movenext
wend
tuijian.close
set tuijian=nothing

elseif Scroll=2 then
set tuijian=server.CreateObject("adodb.recordset")
tuijian.open "select top "&Article&" * from [Products] where tuijian=1 order by id desc",conn,1,1
if tuijian.eof and tuijian.bof then
response.Write("<div class=""onxx"">暂无推荐产品记录</div>")
end if
while not tuijian.eof
if html=0 then
url=""&Dir&"Products/"&ShowProducts&""&Separated&""&tuijian("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Products/ShowProducts.asp?id="&tuijian("id")&""
end if
response.Write("<div class=""tui_box""><a href="""&url&""" title="""&tuijian("title")&"""><img src="""&tuijian("img")&""" width=""213"" height=""130"" alt="""&tuijian("title")&""" onerror=""this.src='images/nologo.jpg'""/><p>"&stvalue(tuijian("title"),24)&"</p></a></div>")
tuijian.movenext
wend
tuijian.close
set tuijian=nothing
end if
end function

Sub Scroll_id()
if Scroll=1 then
response.Write""&zych_Nvachannel(4)&" CAES"
elseif Scroll=2 then
response.Write""&zych_Nvachannel(6)&" PRODUCT"
end if
End Sub


function zych_news_hot(Article)'热点新闻调用参数
set newshot=server.CreateObject("adodb.recordset")
newshot.open "select top "&Article&" * from news order by hit desc",conn,1,1
if newshot.eof and newshot.bof then
response.Write("<div class=""onxx"">暂无暂无热点新闻!</div>")
end if
while not newshot.eof
if html=0 then
response.Write("<li><a href="""&Dir&"News/"&ShowNews&"_"&newshot("id")&"."&HTMLName&""" title="""& newshot("title")&""">"&stvalue(newshot("title"),34)&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"News/ShowNews.asp?id="&newshot("id")&""" title="""& newshot("title")&""">"&stvalue(newshot("title"),34)&"</a></li>")
end if
newshot.movenext
wend
newshot.close
set newshot=nothing
end function

function zych_about_announcement(Article)'首页公告调用参数2
set newshot=server.CreateObject("adodb.recordset")
newshot.open "select top "&Article&" * from announcement order by px_id desc",conn,1,1
if newshot.eof and newshot.bof then
response.Write("<div class=""onxx"">暂无公告</div>")
end if
while not newshot.eof
if html=0 then
response.Write("<a href="""&Dir&"Dynamic/"&announcement&""&Separated&""&newshot("id")&"."&HTMLName&""" title="""& newshot("title") &""">"&stvalue(newshot("title"),40)&"</a><br/>"&vbcrlf)
elseif html=1 then
response.Write("<a href="""&Dir&"Dynamic/?id=" & newshot("id") & """ title="""& newshot("title") &""">" &stvalue(newshot("title"),40)&"</a><br/>"&vbcrlf)
end if
newshot.movenext
wend
newshot.close
set newshot=nothing
end function


function zych_index_flash()'网站主页调用参数
set indexflash=server.CreateObject("adodb.recordset")
indexflash.open "select * from indexflash order by px_id asc",conn,1,1
if indexflash.eof and indexflash.bof then
response.Write("<div class=""onxx"">暂无网站主页图片记录</div>")
end if
while not indexflash.eof
response.Write("<li><a href="""&indexflash("link")&""" target=""_blank""><img src="""&indexflash("img")&""" width=""990px"" height=""400"" alt="""& indexflash("title") &""" /></a></li>")&vbCrLf
indexflash.movenext
wend
indexflash.close
set indexflash=nothing
end function


function zych_index_project()'首页服务项目
set rs=server.createobject("adodb.recordset") 
if BigClassID="" then
exec="select  * from project order by px_id asc"
end if
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div class=""onxx"">暂无服务项目</div> "
end if
do while not rs.eof
if html=0 then
url=""&Dir&"Service/"&Showproject&""&Separated&""&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Service/Showproject.asp?id="&rs("id")&""
end if
response.Write "<div class=""zych_project_box"">"
response.Write "<div class=""leftimg"">"
response.Write "<a href="""&url&""" rel=""lightbox[roadtrip]"" title="""&rs("name")&"""><img src="""&rs("img")&""" alt="""&rs("name")&""" width=""220px"" height=""120px"" onerror=""this.src='images/nologo.jpg'"" /></a></div>"
response.Write "<div class=""righttxt""><a href="""&url&""" title="""&rs("name")&""" class=""title"">"&stvalue(rs("name"),30)&"</a>"
response.Write "<a class=""sm"">"&rs("shuomin")&"</a></div></div>"
rs.movenext
Loop
rs.close
Set rs=nothing
end function


function zych_project_list()'新闻分类调用参数
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from project order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<div class=""onxx"">暂无记录</div>")
end if
while not rs.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Service/"&Showproject&""&Separated&""&rs("id")&"."&HTMLName&""" title="""&rs("name")&""">"&rs("name")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Service/Showproject.asp?id="&rs("id")&""" title="""&rs("name")&""">"&rs("name")&"</a></li>")
end if
rs.movenext
wend
rs.close
set rs=nothing
end function
%>

<%
function zych_link_text()'文字友情链接调用参数
dim link
dim strHtml
set link=server.CreateObject("adodb.recordset")
link.open "select * from link where lx=1 order by px_id asc",conn,1,1
if link.eof and link.bof then
response.Write("<a>暂无链接..</a> ")
end if
while not link.eof
strHtml = strHtml & "<a href="""&link("url")&""" target=""_blank"" title="""&link("title")&""">"&link("title")&"</a> | "
link.movenext
wend
if right(strHtml,3)="|" then
strHtml=left(strHtml,len(strHtml)-3)
end if
link.close
set link=nothing
response.write(strHtml)
end function
%>

<%
function zych_link_img()'图片友情链接调用参数
set link2=server.CreateObject("adodb.recordset")
link2.open "select * from link where lx=2 order by px_id asc",conn,1,1
if link2.eof and link2.bof then
response.Write("<div class=""onxx"">暂无记录</div> !")
end if
while not link2.eof
response.Write("<a href="""&link2("url")&""" title="""&link2("title")&""" target=""_blank""><img src="""&link2("logo")&""" width=""88"" height=""31"" onerror=""this.src='images/nologo.jpg'""/></a>")
link2.movenext
wend
link2.close
set link2=nothing
end function
%>
<%
function zych_qq()'飘浮在线服务QQ
set qq=server.CreateObject("adodb.recordset")
qq.open "select * from zych_qq order by px_id asc",conn,1,1
if qq.eof and qq.bof then
response.Write("<li>暂无QQ</li>")
end if
while not qq.eof
response.Write("<li><span>"&qq("bt")&"</span><a target=""_blank"" href=""http://wpa.qq.com/msgrd?v=3&amp;uin="&qq("qq")&"&amp;site=qq&amp;menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"&qq("qq")&":41"" alt="""&qq("alt")&""" title="""&qq("title")&""" /></a></li>")
qq.movenext
wend
qq.close
set qq=nothing
end function
%>

<%
function zych_contact_qq()'QQ调用参数
dim qq
dim strHtml
set qq=server.CreateObject("adodb.recordset")
qq.open "select top 3 * from zych_qq order by px_id asc",conn,1,1
if qq.eof and qq.bof then
response.Write("暂无QQ")
end if
while not qq.eof
strHtml = strHtml&"<a target=""_blank"" href=""http://wpa.qq.com/msgrd?v=3&amp;uin="&qq("qq")&"&amp;site=qq&amp;menu=yes"">"&qq("qq")&"</a> | "
qq.movenext
wend
if right(strHtml,3)=" | " then
strHtml=left(strHtml,len(strHtml)-3)
end if
qq.close
set qq=nothing
response.write(strHtml)
end function

set rs=server.createobject("adodb.recordset") 
exec="select * from zych_mail" 
rs.open exec,conn,1,1 
E_Server = rs("smtp") ''发件服务器 
E_ServerUser =rs("mailadmin") ''登录用户名 
E_ServerPass =rs("mailpassword") ''登录密码 
E_SendManMail =rs("mailurl") ''发件人邮件地址 
E_SendManName =rs("aajj") ''发件人姓名
E_mail=rs("fmail")
rs.close
set rs=nothing

Sub Jmail(Email,Topic,Mailbody) 
On Error Resume Next 
Dim JMail 
Set JMail = Server.CreateObject("JMail.Message") 
JMail.silent=true 
JMail.Logging = True 
JMail.Charset = "gb2312" 
If Not(E_ServerUser = "" Or E_ServerPass = "") Then 
JMail.MailServerUserName = E_ServerUser 
JMail.MailServerPassword = E_ServerPass 
End If 
JMail.ContentType = "text/html" 
JMail.Priority = 1 
JMail.From = E_SendManMail 
JMail.FromName = E_SendManName 
JMail.AddRecipient Email 
JMail.Subject = Topic 
JMail.Body = Mailbody 
JMail.Send (E_Server) 
Set JMail = Nothing 
SendMail = "OK" 
If Err Then SendMail = "False" 
End Sub 

Sub Cdonts(Email,Topic,Mailbody) 
On Error Resume Next 
Dim ObjCDOMail 
Set ObjCDOMail = Server.CreateObject("CDONTS.NewMail") 
ObjCDOMail.From = E_SendManMail 
ObjCDOMail.To = Email 
ObjCDOMail.Subject = Topic 
ObjCDOMail.BodyFormat = 0 
ObjCDOMail.MailFormat = 0 
ObjCDOMail.Body = Mailbody 
ObjCDOMail.Send 
Set ObjCDOMail = Nothing 
SendMail = "OK" 
If Err Then SendMail = "False" 
End Sub 

Sub Aspemail(Email,Topic,Mailbody) 
On Error Resume Next 
Dim Mailer 
Set Mailer = Server.CreateObject("Persits.MailSender") 
Mailer.Charset = "gb2312" 
Mailer.IsHTML = True 
Mailer.username = E_ServerUser 
Mailer.password = E_ServerPass 
Mailer.Priority = 1 
Mailer.Host = E_Server 
Mailer.Port = 25 
Mailer.From = E_SendManMail 
Mailer.FromName = E_SendManName 
Mailer.AddAddress Email,Email 
Mailer.Subject = Topic 
Mailer.Body = Mailbody 
Mailer.Send 
SendMail = "OK" 
If Err Then SendMail = "False" 
End Sub 
dim SendMail 
Sub SendEmail(Mailto,Subject,HtmlCode,SendMode) 
if SendMode="" then SendMode="Jmail" 
if SendMode="Jmail" then 
Jmail MailTo,Subject,HtmlCode 
elseif SendMode="Cdonts" then 
Cdonts MailTo,Subject,HtmlCode 
elseif SendMode="Aspemail" then 
Aspemail MailTo,Subject,HtmlCode 
end if 
End Sub 
%>