	<%
'��ʾ������Ϣ
sub strA(str1)
      Response.Write("<script language=""JavaScript"">alert("""&str1&""");history.go(-1);</script>")
	  response.End()
end sub	  	

'�ɹ���ʾ��Ϣ
sub strB(str2,url)
      Response.Write("<script language=""JavaScript"">alert("""&str2&""");window.location='"&url&"';</script>")
	  response.End()
end sub	   

'ҳ���Զ���ת
sub AutoJump(str1,url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
	Response.Write("<br/>&nbsp;&nbsp;������ת...<br/>")
	Response.Write("<br/>&nbsp;&nbsp;ҳ��û���Զ���ת<a href="&url&">�������</a><br/>")
	Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
end sub


'��ҳ�ӳ���
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
        Response.Write ("<a>��ҳ</a>")        
    else        
        Response.Write("<A href="&action&"?"&temp&"Page=1>��ҳ</A>")
    end if
if s1>=3 then response.write "<a>..</a>"
for i=s1 to s2
   if i=page then
     response.write "<span>��"&i&"ҳ</span>"
   else
     response.write "<a href="&action&"?"&temp&"Page="&i&">��"&i&"ҳ</a>"
   end if
next
if s2<maxpage then response.write "<a>..</a>"
    Response.Write("<a>ҳ��:"&page&"/"&maxpage&"ҳ</a>")
    Response.Write("<a>��"&iCount&"����¼</a>")
    if page>=maxpage then
        Response.Write ("<a>βҳ</a>")            
    else
        Response.Write("<A href="&action&"?"&temp&"Page="&maxpage&">βҳ</A>")            
    end if
Response.Write("</span>")
'������һҳ��һҳ����
End Sub


sub Left1(str1,url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
	Response.Write("<br/>&nbsp;&nbsp;������ת...<br/>")
	Response.Write("<br/>&nbsp;&nbsp;ҳ��û���Զ���ת<a href="&url&">�������</a><br/>")
	Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
end sub
%>
