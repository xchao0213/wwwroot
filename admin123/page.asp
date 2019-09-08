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
action = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME")
query = Split(Request.ServerVariables("QUERY_STRING"), "&")
  For Each x In query
	  a = Split(x, "=")
	  If StrComp(a(0), "page", vbTextCompare) <> 0 Then
		  temp = temp & a(0) & "=" & a(1) & "&"
	  End If
  Next
Response.Write("<span class=""list_page"">")
if page>3 then s1=page-3 else s1=1
if page<maxpage-3 then s2=page+3 else s2=maxpage
 if page<=1 then
        Response.Write ("<a>首页</a>")        
    else        
        Response.Write("<A href="&action&"?"&temp&"Page=1>首页</A>")
    end if
if s1>=3 then response.write "<a>..</a>"
for i=s1 to s2
   if i=page then
     response.write "<span>第"&i&"页</span>"
   else
     response.write "<a href="&action&"?"&temp&"Page="&i&">第"&i&"页</a>"
   end if
next
if s2<maxpage then response.write "<a>..</a>"
    Response.Write("<a>页次:"&page&"/"&maxpage&"页</a>")
    Response.Write("<a>共"&iCount&"条记录</a>")
    if page>=maxpage then
        Response.Write ("<a>尾页</a>")            
    else
        Response.Write("<A href="&action&"?"&temp&"Page="&maxpage&">尾页</A>")            
    end if
Response.Write("</span>")
'生成上一页下一页链接
End Sub


sub Left1(str1,url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
	Response.Write("<br/>&nbsp;&nbsp;正在跳转...<br/>")
	Response.Write("<br/>&nbsp;&nbsp;页面没有自动跳转<a href="&url&">【点这里】</a><br/>")
	Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
end sub
%>
