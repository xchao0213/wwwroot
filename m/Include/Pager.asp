<%
'��ҳ�ӳ���
Sub PageControl(iCount,pagecount,page)
'������һҳ��һҳ����
    Dim query, a, x, temp
    action = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp&a(0)&"="&a(1)&"&"
        End If
    Next
    if page<=1 then      
        Response.Write ("<a><</a>")
    else        
        Response.Write("<A href="&action&"?"&temp&"Page="&(Page-1)&"> < </A>")
    end if
	for i=1 to maxpage
	if page=i then
        response.write "<span>"&i&"</span>"
	else 
	  response.write "<a href="&action&"?"&temp&"Page="&i&">"&i&"</a>"
	end if
    next
		Response.Write ("<a>��"&maxpage&"ҳ</a>")
    if page>=pagecount then
        Response.Write ("<a>></a>")   
    else
        Response.Write("<A href="&action&"?"&temp&"Page="&(Page+1)&"> > </A> "&vbCrLf)          
    end if
   
End Sub

%>