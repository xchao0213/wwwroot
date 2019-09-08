<div class="logobox"><div class="logo"><a href="<%=dir%>"><img alt="<%=zych_home%>" src="<%=zych_logo%>" height="50" onerror="this.src='<%=dir%>images/nologo.jpg'"></a></div><a href="javascript:history.go(-1)" class="fan"><span>·µ»Ø</span></a></div>
<header>

<ul class="navigation">
<%set rsn=server.CreateObject("adodb.recordset")
rsn.open "select * from menu where yc=1 order by px_id asc",conn,1,1
while not rsn.eof
i=i+1
if "m/"&rsn("url")=wurl  then  active=" class=""active"""  else active=""
if Instr(rsn("url"),"http")>0 then navurl=rsn("url") else navurl=dir&"m/"&rsn("url")
' if i=1 then navurl=rsn("url") else navurl="a" 
response.Write("<li"&active&"><a href="""&navurl&""">"&rsn("title")&"</a></li>"&vbcrlf)

rsn.movenext
wend
rsn.close
set rsn=nothing %>      
</ul>
<!--Search starts-->

</header>

