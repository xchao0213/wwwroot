<%
	Dim GetFlag Rem(�ύ��ʽ)
	Dim ErrorSql Rem(�Ƿ��ַ�) 
	Dim RequestKey Rem(�ύ����)
	Dim ForI Rem(ѭ�����)
	ErrorSql = "'~;~and~(~)~exec~update~count~*~%~chr~mid~master~truncate~char~declare" Rem(ÿ�������ַ����ߴ�����ʹ�ð�� "~" ��)
	ErrorSql = split(ErrorSql,"~")
	If Request.ServerVariables("REQUEST_METHOD")="GET" Then
	GetFlag=True
	Else
	GetFlag=False
	End If
	If GetFlag Then
	  For Each RequestKey In Request.QueryString
	   For ForI=0 To Ubound(ErrorSql)
		If Instr(LCase(Request.QueryString(RequestKey)),ErrorSql(ForI))<>0 Then
		response.write "<script>alert(""����:�벻Ҫ���ԷǷ�ע�룡��"");history.go(-1);</script>"
		Response.End
		End If
	   Next
	  Next 
	Else
	  For Each RequestKey In Request.Form
	   For ForI=0 To Ubound(ErrorSql)
		If Instr(LCase(Request.Form(RequestKey)),ErrorSql(ForI))<>0 Then
		response.write "<script>alert(""����:�벻Ҫ���ԷǷ�ע�룡��"");history.go(-1);</script>"
		Response.End
		End If
	   Next
	  Next
	End If
%>

