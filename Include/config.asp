<!--#include file="Inc.Asp" -->
<!--#include file="CommonFun.asp" -->
<%az=request.ServerVariables("SERVER_NAME")
webd=ubound(split(az,"."))
webq=Replace(az,".","")
If webd>=3 or az="localhost" or isnumeric(webq) Then 
Response.Write ""
else
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(""&Dir&"Include/NoInstall.txt")) Then
Response.Write ""
else
response.Redirect ""&Dir&"Include/install.asp"
end if
set fso=nothing
end if
zych_c=Chr("32")&Chr("45")&Chr("80")&Chr("111")&Chr("119")&Chr("101")&Chr("114")&Chr("101")&Chr("100")&Chr("32")&Chr("98")&Chr("121")&Chr("32")&Chr("122")&Chr("121")&Chr("99")&Chr("104")&Chr("114")&Chr("46")&Chr("99")&Chr("111")&Chr("109")&""
zych_c=""
set config=server.createobject("adodb.recordset") 
exec="select * from config" 
config.open exec,conn,1,1 
if config("on")=1 then
response.redirect ""&Dir&"webClose.asp"
response.End
else
zych_home=""&config("title")&zych_c&""
zych_keywords=""&config("keywords")&""
zych_description=""&config("description")&""
zych_url=""&config("url")&""
zych_logo=""&config("logo")&""
zych_logoiwidth=""&config("logoiwidth")&""
zych_logoheight=""&config("logoheight")&""
zych_name=""&config("name")&""
zych_css=""&config("css")&""
zych_mail=""&config("mail")&""
zych_Phone=""&config("Phone")&""
zych_tel=""&config("tel")&""
zych_fax=""&config("fax")&""
zych_Postal=""&config("Postal")&""
zych_dz=""&config("dz")&""
zych_copyright=""&config("copyright")&""
zych_logoright=""&config("logoright")&""
zych_daohangnr=""&config("daohangnr")&""
zych_booksh=""&config("booksh")&""
zych_beian="<a href=""http://www.miibeian.gov.cn""  target=""_blank"">"&config("beian")&"</a>"
zych_video_sz=""&config("video_sz")&""
zych_video_wz=""&config("video_wz")&""
zych_zztj=""&config("zztj")&""
zych_foot_left=""&config("foot_left")&"" 
end if
config.close
set config=nothing

function weburl()'��վURL
If Request.ServerVariables("SERVER_PORT") = "80" Then
GetSiteUrl = "http://" & Request.ServerVariables("server_name")
Else
GetSiteUrl = "http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("SERVER_PORT")
End If
end function

function zych_Visits()
sql="update config set js=js+1 where js is not null" 
conn.execute(sql) 
response.Write(""&config("bq")&" ��������"&config("js")&"")
end function
call zych_Nav()

wz=request.ServerVariables("HTTP_HOST")&request.ServerVariables("URL")
wurl=right(left(wz,instrrev(wz,"/")),(len(left(wz,instrrev(wz,"/")))-len(request.ServerVariables("HTTP_HOST")))-1)
function zych_class(wurl)'�˵��������ò���
set nav=server.CreateObject("adodb.recordset")
nav.open "select * from menu ",conn,1,1
if nav.eof and nav.bof then
response.Write ""
else
while not Nav.eof

	if dir&nav("url")="/"&wurl then 
		response.Write nav("title")&""
 	end if
Nav.movenext
wend	
end if
nav.close
set nav=nothing
end function

function zych_pd_class(wurl)'�˵��������ò���
set nav=server.CreateObject("adodb.recordset")
nav.open "select * from menu ",conn,1,1
if nav.eof and nav.bof then
response.Write ""
else
while not Nav.eof
	if dir&nav("url")="/"&wurl then 
		response.Write nav("title")&" <i>"&nav("en_title")&"</i>"
 	end if
Nav.movenext
wend	
end if
nav.close
set nav=nothing
end function

function zych_Navigation()'�˵��������ò���
set Nav=server.CreateObject("adodb.recordset")
Nav.open "select * from menu where yc=1 order by px_id asc",conn,1,1
while not Nav.eof
if Nav("url")=wurl then no=" class=""selected""" else no=""
if Instr(Nav("url"),"http")>0 then navurl=Nav("url") else navurl=dir&Nav("url")
Response.Write("<li><a href="""&navurl&""" "&no&" target="""&Nav("openfs")&""">"&Nav("title")&"</a>")
'�����˵���ʼ
  set nb=server.CreateObject("adodb.recordset")
  nb.open "select * from menu_nav where ssfl="&nav("id")&" and  yc=1 order by px_id asc",conn,1,1
  if nb.eof and nb.bof then
  else
  Response.Write("<ul class=""submenu"">"&vbcrlf)
  while not nb.eof
  Response.Write("<li><a href="""&dir&nb("url")&""" title="""&nb("title")&""">"&nb("title")&"</a></li>"&vbcrlf)
  nb.movenext
  wend
  Response.Write("</ul>"&vbcrlf)
  end if
  nb.close
  set nb=nothing
Response.Write("</li>"&vbcrlf)
Nav.movenext
wend
Nav.close
set Nav=nothing
end function

function zych_Navigation_top()'��վ�����˵��������ò���
dim Navigation
set nav=server.CreateObject("adodb.recordset")
nav.open "select * from menu where yc=2 order by px_id asc",conn,1,1
while not nav.eof
if Instr(Nav("url"),"http://")>0 then navurl=Nav("url") else navurl=dir&Nav("url")
response.Write("<a href="""&navurl&""" target="""&nav("openfs")&""">"&nav("title")&"</a>"&vbcrlf)
nav.movenext
wend
nav.close
set nav=nothing
end function

function zych_about_list()'��ҳ�б��ֵ��ò���
set aboutlist=server.CreateObject("adodb.recordset")
aboutlist.open "select * from about where xianshi=1 order by px_id asc",conn,1,1
if aboutlist.eof and aboutlist.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not aboutlist.eof
if html=0 then
response.Write("<li><a href="""&Dir&"About/"&about&Separated&aboutlist("id")&"."&HTMLName&""" title="""&aboutlist("title")&""">"  & aboutlist("title") & "</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"About/?id="&aboutlist("id")&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a></li>")
end if
aboutlist.movenext
wend
aboutlist.close
set aboutlist=nothing
end function


function zych_about_list_pd()'��ҳ�б��ֵ��ò���
set aboutlist=server.CreateObject("adodb.recordset")
aboutlist.open "select * from about where xianshi=1 order by px_id asc",conn,1,1
if aboutlist.eof and aboutlist.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not aboutlist.eof
if html=0 then
response.Write("<li><a href="""&Dir&"About/"&about&Separated&aboutlist("id")&"."&HTMLName&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"About/?id="&aboutlist("id")&""" title="""&aboutlist("title")&""">"&aboutlist("title")&"</a></li>")
end if
aboutlist.movenext
wend
aboutlist.close
set aboutlist=nothing
end function


function zych_News_fl_list()'���ŷ�����ò���
set rsfl=server.CreateObject("adodb.recordset")
rsfl.open "select * from news_fl order by px_id asc",conn,1,1
if rsfl.eof and rsfl.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not rsfl.eof
if html=0 then
response.Write("<li><a href="""&Dir&"News/"&Newsfl&Separated&rsfl("id")&"."&HTMLName&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"News/Newslist.asp?id="&rsfl("id")&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
end if
rsfl.movenext
wend
rsfl.close
set rsfl=nothing
end function

function zych_Culture_fl_list()'�Ļ�������ò���
set rsfl=server.CreateObject("adodb.recordset")
rsfl.open "select * from culture_fl order by px_id asc",conn,1,1
if rsfl.eof and rsfl.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not rsfl.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Culture/"&Newsfl&Separated&rsfl("id")&"."&HTMLName&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"/Culture/Culturelist.asp?id="&rsfl("id")&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
end if
rsfl.movenext
wend
rsfl.close
set rsfl=nothing
end function

function zych_News_fl_list()'԰��������ò���
set rsfl=server.CreateObject("adodb.recordset")
rsfl.open "select * from fengfa_fl order by px_id asc",conn,1,1
if rsfl.eof and rsfl.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not rsfl.eof
if html=0 then
response.Write("<li><a href="""&Dir&"News/"&Newsfl&Separated&rsfl("id")&"."&HTMLName&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Fengfa/Fengfalist.asp?id="&rsfl("id")&""" title="""&rsfl("title")&""">"&rsfl("title")&"</a></li>")
end if
rsfl.movenext
wend
rsfl.close
set rsfl=nothing
end function

%>

<%function zych_Flash()'Flash�õƵ��ò���%>
<SCRIPT language=JavaScript type=text/javascript>
var swf_width='990';
var swf_height='200';
var configtg='0xffffff:������ɫ|2:����λ��|0x000000:���ֱ�����ɫ|10:���ֱ���͸����|0xffffff:����������ɫ|0x4f6898:����Ĭ����ɫ|0x000033:������ǰ��ɫ|6:�Զ�����ʱ��|3:ͼƬ����Ч��|1:�Ƿ���ʾ��ť|_blank:���´���';
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
response.write "var files='"&vbcrlf
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

<%function zych_index_video()'��ҳ��Ƶ���ò���
set rs=server.createobject("adodb.recordset") 
exec="select top 1 *  from [video] where xianshi=1 order by px_id asc"
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div style=""padding:10px"">�޴���Ƶ...</div>"
else
if suffix(rs("url"))=".flv" then
response.Write "<div class=""flv"">"%>
<script type="text/javascript">
var swf_width=250
var swf_height=200
texts='<%=rs("title")%>'
var files='<%=rs("url")%>'
var config='<%=vid_pay%>:�Զ�����|1:��������|100:Ĭ������|0:������λ��|2:��������ʾ|0x000033:������ɫ|60:����͸����|0x66ff00:������ɫ|0xffffff:ͼ����ɫ|0xffffff:������ɫ|<%=vid_wz%>:����:logo��ַ|:����swf��ַ'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="<%=dir%>images/vcastr2.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'">');
document.write('<embed src="<%=dir%>images/vcastr2.swf" wmode="opaque" FlashVars="vcastr_file='+files+'&vcastr_title='+texts+'&vcastr_config='+config+'" menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
</SCRIPT>
<%else
newstr=split(rs("url"),"/id_")'��ַ�ָ�
urlid=split(newstr(1),".")
if vid_pay=1 then AutoPlay="true" else AutoPlay="false"
response.Write "<div class=""swf"">"
Response.Write"<embed src=""http://static.youku.com/v1.0.0149/v/swf/loader.swf?VideoIDS="&urlid(0)&"&winType=adshow&isAutoPlay="&AutoPlay&""" quality=""high"" width=""250"" height=""200"" type=""application/x-shockwave-flash""></embed>"
end if
rs.close
set rs=nothing
response.Write "</div>"
end if
end function%>

<%
function zych_download_fl()'���ط�����ò���
set downloadclas=server.CreateObject("adodb.recordset")
downloadclas.open "select * from download_fl order by px_id asc",conn,1,1
if downloadclas.eof and downloadclas.bof then
response.Write("&nbsp;���޼�¼ !")
end if
while not downloadclas.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Download/"&Downloadfl&Separated&"" & downloadclas("id") & "."&HTMLName&""" title=""" & downloadclas("title") &""">"&downloadclas("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Download/?id=" & downloadclas("id")&""" title=""" &downloadclas("title")&""">"& downloadclas("title")&"</a></li>")
end if
downloadclas.movenext
wend
downloadclas.close
set downloadclas=nothing
end function

function zych_case_class()'����������ò���
set anliclas=server.CreateObject("adodb.recordset")
anliclas.open "select * from case_class order by px_id asc",conn,1,1
if anliclas.eof and anliclas.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not anliclas.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Case/"&Casefl&Separated&"" & anliclas("id") & "."&HTMLName&""" title=""" & anliclas("title") &""">"  & anliclas("title") & "</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Case/?id=" & anliclas("id")&""" title=""" &anliclas("title")&""">"& anliclas("title")&"</a></li>")
end if
anliclas.movenext
wend
anliclas.close
set anliclas=nothing
end function


function zych_Products_class()'��Ʒ����
set rs1=server.CreateObject("adodb.recordset")
rs1.open "select * from [bigClass] order by px_id asc",conn,1,1
if rs1.eof and rs1.bof then
response.Write("<div class=""p_menu_1""><a>���޼�¼</a></div> ")
end if
while not rs1.eof
if html=0 then
url=""&Dir&"Products/"&Productsfl&""&Separated&""&rs1("BigClassID")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Products/?BigClassID="&rs1("BigClassID")&""
end if
response.write("<div class=""p_menu_1""><a href="""&url&""" title="&rs1("BigClassName")&">"&rs1("BigClassName")&"</a></div>"&vbcrlf)
'�������࿪ʼ
set rs2=server.createobject("adodb.recordset") 
exec="select * from [SmallClass] where BigClassID="&rs1("BigClassID")&" order by px_id asc  " 
rs2.open exec,conn,1,1
if rs2.eof and rs2.bof then
response.Write("")
else
Response.Write("<div class=""p_menu_2"">") 
do while not rs2.eof
if html=0 then
url2=""&Dir&"Products/"&Productsfl&""&Separated&""&rs2("BigClassID")&""&Separated&""&rs2("SmallClassID")&"."&HTMLName&""
elseif html=1 then
url2=""&Dir&"Products/?BigClassID="&rs2("BigClassID")&"&amp;SmallClassID="&rs2("SmallClassID")&""
end if
response.write("<li><a href="""&url2&""" title="&rs1("BigClassName")&" class=""on2"">"&rs2("SmallClassName")&"</a></li>")
rs2.movenext
loop
Response.Write("</div>")
end if
rs2.close
set rs2=nothing
rs1.movenext
wend
rs1.close
set rs1=nothing
end function


function zych_photo_class()'��������ò���
set photo=server.CreateObject("adodb.recordset")
photo.open "select * from photo_class order by px_id asc",conn,1,1
if photo.eof and photo.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not photo.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Photo/"&photofl&Separated&"" & photo("id") & "."&HTMLName&""" title=""" & photo("title") &""">"&photo("title")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Photo/?id=" & photo("id") & """ title=""" & photo("title") &""">"  & photo("title") & "</a></li>")
end if
photo.movenext
wend
photo.close
set photo=nothing
end function

function zych_video_class()'��Ƶ������ò���
set video=server.CreateObject("adodb.recordset")
video.open "select * from video_class order by px_id asc",conn,1,1
if video.eof and video.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not video.eof
if html=0 then
response.Write("<li><a href="""&Dir&"video/"&Videofl&Separated&"" & video("id") & "."&HTMLName&""" title=""" & video("title") &""">"  & video("title") & "</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"video/?id=" & video("id") & """ title=""" & video("title") &""">"  & video("title") & "</a></li>")
end if
video.movenext
wend
video.close
set video=nothing
end function

function zych_index_case(Article)'��ҳ�Ƽ�����ͼ�ĵ��ò���
if Scroll=1 then
set tuijian=server.CreateObject("adodb.recordset")
tuijian.open "select top "&Article&" * from [case] where tuijian=1 order by id desc",conn,1,1
if tuijian.eof and tuijian.bof then
response.Write("<div class=""onxx"">�����Ƽ�������¼</div>")
end if
while not tuijian.eof
if html=0 then
url=""&Dir&"Case/"&showcase&Separated&tuijian("id")&"."&HTMLName&""
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
response.Write("<div class=""onxx"">�����Ƽ���Ʒ��¼</div>")
end if
while not tuijian.eof
if html=0 then
url=""&Dir&"Products/"&ShowProducts&Separated&tuijian("id")&"."&HTMLName&""
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

function zych_news_hot(Article)'�ȵ����ŵ��ò���
set newshot=server.CreateObject("adodb.recordset")
newshot.open "select top "&Article&" * from news order by hit desc",conn,1,1
if newshot.eof and newshot.bof then
response.Write("<div class=""onxx"">���������ȵ�����!</div>")
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

function zych_job_list()'ְλ�б�
set rsj=server.CreateObject("adodb.recordset")
rsj.open "select top 10 * from job order by id desc",conn,1,1
if rsj.eof and rsj.bof then
response.Write("&nbsp;���޼�¼ !")
end if
while not rsj.eof
if html<>0 then
response.Write("<li><a href="""&Dir&"job/Showjob.asp?id="&rsj("id")&""" title="""&rsj("title")&""">"&stvalue(rsj("title"),34)&"</a></li>")
else
response.Write("<li><a href="""&dir&"job/"&Showjob&""&Separated&""&rsj("id")&"."&HTMLName&""" title="""&rsj("title")&""">"&stvalue(rsj("title"),34)&"</a></li>")
end if
rsj.movenext
wend
rsj.close
set rsj=nothing
end function


function zych_index_flash()'��վ��ҳ���ò���
set rsf=server.CreateObject("adodb.recordset")
rsf.open "select * from indexflash order by px_id asc",conn,1,1
if rsf.eof and rsf.bof then
response.Write("<div class=""onxx"">������վ��ҳͼƬ��¼</div>")
end if
while not rsf.eof
response.Write("<li src=""url("&rsf("img")&")"" style=""background:"&rsf("coler")&" no-repeat center top""><a href="""&rsf("link")&""" target=""_blank""></a></li>")&vbCrLf
rsf.movenext
wend
rsf.close
set rsf=nothing
end function


function zych_Service_list()'������Ŀ�б�
set rs=server.createobject("adodb.recordset") 
if BigClassID="" then
exec="select  * from project order by px_id asc"
end if
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div class=""onxx"">���޷�����Ŀ</div> "
end if
do while not rs.eof
if html=0 then
url=""&Dir&"Service/"&Showproject&Separated&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Service/Showproject.asp?id="&rs("id")&""
end if
response.Write "<li><a href="""&url&"""  title="""&rs("name")&""">"&rs("name")&"</a></li>"
rs.movenext
Loop
rs.close
Set rs=nothing
end function

function zych_Service_list()'�쵼�����б�
set rs=server.createobject("adodb.recordset") 
if BigClassID="" then
exec="select  * from leader order by px_id asc"
end if
rs.open exec,conn,1,1
if rs.eof then
response.Write "<div class=""onxx"">���޷�����Ŀ</div> "
end if
do while not rs.eof
if html=0 then
url=""&Dir&"Service/"&Showproject&Separated&rs("id")&"."&HTMLName&""
elseif html=1 then
url=""&Dir&"Service/ShowLeader.asp?id="&rs("id")&""
end if
response.Write "<li><a href="""&url&"""  title="""&rs("name")&""">"&rs("name")&"</a></li>"
rs.movenext
Loop
rs.close
Set rs=nothing
end function

function zych_login()'��Ա��½���ò���
if session("username")<>"" then 
set rss=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rss.open exec,conn,1,1 
response.write("<a href="""&Dir&"User/?action=admin"">["&rss("useradmin")&"]</a> ��ӭ��!<a href="""&Dir&"User/LoginOut.asp"">�˳�</a>")
else
response.write("<a href="""&Dir&"User/login.asp"">[��Ա��½]</a>")
end if
end function

function zych_user_nav()'��Ա���ĵ���
Response.Write("<li><a href="""&Dir&"User/?action=admin"">��Ա��ҳ</a></li>")
Response.Write("<li><a href="""&Dir&"User/?action=Edit"">�ҵ�����</a></li>")
Response.Write("<li><a href="""&Dir&"User/my_addlist.asp"">�� �� ��</a></li>")
Response.Write("<li><a href="""&Dir&"User/my_orders.asp"">�ҵĶ���</a></li>")
Response.Write("<li><a href="""&Dir&"User/?action=password"">�޸�����</a></li>")
Response.Write("<li><a href="""&Dir&"User/LoginOut.asp"">�˳���½</a></li>")
end function

function zych_project_list()'���ŷ�����ò���
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from project order by px_id asc",conn,1,1
if rs.eof and rs.bof then
response.Write("<div class=""onxx"">���޼�¼</div>")
end if
while not rs.eof
if html=0 then
response.Write("<li><a href="""&Dir&"Service/"&Showproject&Separated&rs("id")&"."&HTMLName&""" title="""&rs("name")&""">"&rs("name")&"</a></li>")
elseif html=1 then
response.Write("<li><a href="""&Dir&"Service/Showproject.asp?id="&rs("id")&""" title="""&rs("name")&""">"&rs("name")&"</a></li>")
end if
rs.movenext
wend
rs.close
set rs=nothing
end function

function zych_link_text()'�����������ӵ��ò���
dim link
dim strHtml
set link=server.CreateObject("adodb.recordset")
link.open "select * from link where lx=1 order by px_id asc",conn,1,1
if link.eof and link.bof then
response.Write("��������...")
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

function zych_link_img()'ͼƬ�������ӵ��ò���
set link2=server.CreateObject("adodb.recordset")
link2.open "select * from link where lx=2 order by px_id asc",conn,1,1
if link2.eof and link2.bof then
response.Write ""
end if
while not link2.eof
response.Write("<a href="""&link2("url")&""" title="""&link2("title")&""" class=""pic"" target=""_blank""><img src="""&link2("logo")&"""  height=""35"" onerror=""this.src='images/nologo.jpg'""/></a>")
link2.movenext
wend
link2.close
set link2=nothing
end function

function zych_qq()'Ʈ�����߷���QQ
set qq=server.CreateObject("adodb.recordset")
qq.open "select * from zych_qq order by px_id asc",conn,1,1
if qq.eof and qq.bof then
response.Write("<li>����QQ</li>")
end if
while not qq.eof
response.Write("<li><span>"&qq("bt")&"</span><a target=""_blank"" href=""http://wpa.qq.com/msgrd?v=3&amp;uin="&qq("qq")&"&amp;site=qq&amp;menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"&qq("qq")&":41"" alt="""&qq("alt")&""" title="""&qq("title")&""" /></a></li>")
qq.movenext
wend
qq.close
set qq=nothing
end function

function zych_contact_qq()'�����������ӵ��ò���
dim qq
dim strHtml
set qq=server.CreateObject("adodb.recordset")
qq.open "select top 3 * from zych_qq order by px_id asc",conn,1,1
if qq.eof and qq.bof then
response.Write("����QQ")
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

if session("username")<>"" then
set rss=server.createobject("adodb.recordset") 
exec="select * from [user] where id="&session("username")&"  " 
rss.open exec,conn,1,1 
userid=rss("id")
useradmin=rss("useradmin")
username=rss("zsname")
usergs=rss("gsname")
useradd=rss("gsadd")
usersj=rss("sj")
usertel=rss("tel")
userqq=rss("qq")
rss.close
set rss=nothing
end if

Private Sub add_echo(text,ats,aurl,bts,burl)
Response.Write("<div id=""divts""><h2>"&text&"</h2>")
Response.Write("<div class=""ts"">")
Response.Write("<a href="""&aurl&""" class=""tl"">"&ats&"</a>")
Response.Write("<a href="""&burl&""" class=""tr"">"&bts&"</a></div>")
Response.Write("</div>")
End Sub
set rs=server.createobject("adodb.recordset") 
exec="select * from zych_mail" 
rs.open exec,conn,1,1 
E_Server = rs("smtp") ''���������� 
E_ServerUser =rs("mailadmin") ''��¼�û��� 
E_ServerPass =rs("mailpassword") ''��¼���� 
E_SendManMail =rs("mailurl") ''�������ʼ���ַ 
E_SendManName =rs("aajj") ''����������
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
Public Function getTime()
	getTime = Right(getStrNow,12)
End Function
'��ȡʱ���ַ���, ��ʽYYYYMMDDhhmiss
Public Function getStrNow()
	strNow = Now()
	strNow = Year(strNow) & Right(("00" & Month(strNow)),2) & Right(("00" & Day(strNow)),2) & Right(("00" & Hour(strNow)),2) & Right(("00" &  Minute(strNow)),2) & Right(("00" & Second(strNow)),2)
	getStrNow = strNow
End Function

'��ȡ�����,���� [min,max]��Χ����
Public Function getRandNumber(max, min)
	Randomize 
	getRandNumber = CInt((max-min+1)*Rnd()+min) 
End Function

'��ȡ������ֵ��ַ���,����[min,max]��Χ�������ַ���
Public Function getStrRandNumber(max, min)
	randNumber = getRandNumber(max, min)
	getStrRandNumber = CStr(randNumber)
End Function
%>