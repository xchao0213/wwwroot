<%
'HTML ����
Function textHTMLEncode(fString)
    fString = Replace(fString, Chr(62), "&gt;")
    fString = Replace(fString, Chr(60), "&lt;")
    fString = Replace(fString, Chr(32), "&nbsp;")
    fString = Replace(fString, Chr(9), "&nbsp;")
    fString = Replace(fString, Chr(34), "&quot;")
    fString = Replace(fString, Chr(39), "&#39;")
    fString = Replace(fString, Chr(38), "&amp;")
    fString = Replace(fString, Chr(13) & Chr(10), "</P><P>")
    fString = Replace(fString, Chr(13), "<BR>")
    fString = Replace(fString, Chr(10), "<BR>")
    textHTMLEncode = fString
End Function

'HTML ����
Function textHTMLDecode(fString)
    fString = Replace(fString, "&gt;", Chr(62))
    fString = Replace(fString, "&lt;", Chr(60))
    fString = Replace(fString, "&nbsp;", Chr(32))
    fString = Replace(fString, "&nbsp;", Chr(9))
    fString = Replace(fString, "&quot;", Chr(34))
    fString = Replace(fString, "&#39;", Chr(39))
    fString = Replace(fString, "&amp;", Chr(38))
    fString = Replace(fString, "</P><P>", Chr(13) & Chr(10))
    fString = Replace(fString, "<BR>", Chr(10))
    textHTMLDecode = fString
End Function


'=========��ʽ������==================
Function  DateFormat(DateStr,Types)
    Dim DateString
	if IsDate(DateStr) = False then
		DateString = ""
	end if
	Select Case Types
	  Case "1" 
		  DateString = Year(DateStr)&"-"&Month(DateStr)&"-"&Day(DateStr)
	  Case "2"
		  DateString = Year(DateStr)&"."&Month(DateStr)&"."&Day(DateStr)
	  Case "3"
		  DateString = Year(DateStr)&"/"&Month(DateStr)&"/"&Day(DateStr)
	  Case "4"
		  DateString = Month(DateStr)&"/"&Day(DateStr)&"/"&Year(DateStr)
	  Case "5"
		  DateString = Day(DateStr)&"/"&Month(DateStr)&"/"&Year(DateStr)
	  Case "6"
		  DateString = Month(DateStr)&"-"&Day(DateStr)&"-"&Year(DateStr)
	  Case "7"
		  DateString = Month(DateStr)&"."&Day(DateStr)&"."&Year(DateStr)
	  Case "8"
		  DateString = Month(DateStr)&"-"&Day(DateStr)
	  Case "9"
		  DateString = Month(DateStr)&"/"&Day(DateStr)
	  Case "10"
		  DateString = Month(DateStr)&"."&Day(DateStr)
	  Case "11"
		  DateString = Month(DateStr)&"��"&Day(DateStr)&"��"
	  Case "12"
		  DateString = Day(DateStr)&"��"&Hour(DateStr)&"ʱ"
	  case "13"
		  DateString = Day(DateStr)&"��"&Hour(DateStr)&"��"
	  Case "14"
		  DateString = Hour(DateStr)&"ʱ"&Minute(DateStr)&"��"
	  Case "15"
		  DateString = Hour(DateStr)&":"&Minute(DateStr)
	  Case "16"
		  DateString = Year(DateStr)&"��"&Month(DateStr)&"��"&Day(DateStr)&"��"
	  Case Else
	  	  DateString = DateStr
	 End Select
	 DateFormat = DateString
 End Function


'ȥ��html��Ǻ�����������:
Function RemoveHTML( strText ) 
    Dim TAGLIST 
    TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" &_ 
              "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" &_ 
              "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" &_ 
              "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" &_ 
              "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" &_ 
              "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;P;PARAM;PLAINTEXT;" &_ 
              "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" &_ 
              "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;" 

    Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;" 
     
    Dim nPos1 
    Dim nPos2 
    Dim nPos3 
    Dim strResult 
    Dim strTagName 
    Dim bRemove 
    Dim bSearchForBlock 
     
    nPos1 = InStr(strText, "<") 
    Do While nPos1 > 0 
        nPos2 = InStr(nPos1 + 1, strText, ">") 
        If nPos2 > 0 Then 
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1) 
        strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ") 

            nPos3 = InStr(strTagName, " ") 
            If nPos3 > 0 Then 
                strTagName = Left(strTagName, nPos3 - 1) 
            End If 
             
            If Left(strTagName, 1) = "/" Then 
                strTagName = Mid(strTagName, 2) 
                bSearchForBlock = False 
            Else 
                bSearchForBlock = True 
            End If 
             
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then 
                bRemove = True 
                If bSearchForBlock Then 
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then 
                        nPos2 = Len(strText) 
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare) 
                        If nPos3 > 0 Then 
                            nPos3 = InStr(nPos3 + 1, strText, ">") 
                        End If 
                         
                        If nPos3 > 0 Then 
                            nPos2 = nPos3 
                        End If 
                    End If 
                End If 
            Else 
                bRemove = False 
            End If 
             
            If bRemove Then 
                strResult = strResult & Left(strText, nPos1 - 1) 
                strText = Mid(strText, nPos2 + 1) 
            Else 
                strResult = strResult & Left(strText, nPos1) 
                strText = Mid(strText, nPos1 + 1) 
            End If 
        Else 
            strResult = strResult & strText 
            strText = "" 
        End If 
         
        nPos1 = InStr(strText, "<") 
    Loop 
    strResult = strResult & strText 
     
    RemoveHTML = strResult 
End Function 





'*************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ�����...
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function


'*************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ�������...
'*************************************************
function gotTopic1(str,strlen)
	if str="" then
		gotTopic1=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic1=left(str,i)
			exit for
		else
			gotTopic1=str
		end if
	next
	gotTopic1=replace(replace(replace(replace(gotTopic1," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'��ȡ�ļ�
Function ReadtxtFiles(Filename)
   Dim fso, f1, ts, s
   Const ForReading = 1
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set ts = fso.OpenTextFile(Filename, ForReading)
   ReadtxtFiles = ts.ReadAll
   ts.Close
End Function

'�����ļ�
Sub CreateFile(Filename,temp)
 '  Dim fso, tf ,Filename ,temp
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set tf = fso.CreateTextFile(Filename, True)
   ' дһ�У����Ҵ��������ַ���
'   tf.WriteLine("Testing 1, 2, 3.") 
   '���ļ�д���������ַ���        
'   tf.WriteBlankLines(3) 
   'дһ�С�
'   tf.Write ("This is a test.") 
   tf.write temp
   tf.Close
End Sub

'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'pos=InStr(1,"abcdefg","cd") 
'��pos�᷵��3��ʾ���ҵ�����λ��Ϊ�������ַ���ʼ��
'����ǡ����ҡ���ʵ�֣�����������һ�������ܵ�
'ʵ�־��ǰѵ�ǰλ����Ϊ��ʼλ�ü������ҡ�
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'***********************************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub

'********************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'***************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("�й�")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'****************************************************
'��������SendMail
'��  �ã���Jmail��������ʼ�
'��  ����ServerAddress  ----��������ַ
'        AddRecipient  ----�����˵�ַ
'        Subject       ----����
'        Body          ----�ż�����
'        Sender        ----�����˵�ַ
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>û�а�װJMail���</li>"
		err.clear
		exit function
	end if
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=MailServerAddress
	JMail.AddRecipient=AddRecipient
	JMail.Subject=Subject
	JMail.Body=MailBody
	JMail.Sender=Sender
	JMail.From = MailFrom
	JMail.Priority=1
	JMail.Execute 
	Set JMail=nothing 
	if err then 
		SendMail=err.description
		err.clear
	else
		SendMail="OK"
	end if
end function



function getFileExtName(fileName)
    dim pos
    pos=instrrev(filename,".")
    if pos>0 then 
        getFileExtName=mid(fileName,pos+1)
    else
        getFileExtName=""
    end if
end function 

'****************************************************
'��������ReplaceRemoteUrl
'��  �ã�Զ������ͼƬ
'****************************************************

Const sFileExt="jpg|gif|bmp|png"
Function ReplaceRemoteUrl(sHTML, sSaveFilePath, sFileExt)
	 If Not fso.FolderExists(Server.mappath(uploadpath)) Then
	  fso.CreateFolder(Server.mappath(uploadpath))'����Ŀ¼
	 End If
     Dim s_Content
     s_Content = sHTML
     If IsObjInstalled("Microsof"&"t.X"&"MLHTTP") = False then
         ReplaceRemoteUrl = s_Content
         Exit Function
     End If     
     Dim re, RemoteFile, RemoteFileurl,SaveFileName,SaveFileType,arrSaveFileNameS,arrSaveFileName,sSaveFilePaths
     Set re = new RegExp
     re.IgnoreCase = True
     re.Global = True
     re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & sFileExt & ")))"
     Set RemoteFile = re.Execute(s_Content)
     For Each RemoteFileurl in RemoteFile
		 arrSaveFileName = Split(RemoteFileurl,".")
  		 SaveFileType=arrSaveFileName(UBound(arrSaveFileName))
		 RanNum=Int(900*Rnd)+100&Chr(95)&Chr(90)&Chr(89)&Chr(67)&Chr(72)
         arrSaveFileName = Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&ranNum&"."&SaveFileType
  sSaveFilePaths= sSaveFilePath
         SaveFileName = sSaveFilePaths & arrSaveFileName 
		 if SaveRemoteFile(""&SaveFileName&"",""&RemoteFileurl&"") then 
		 'response.Write ""&SaveFileName&" ͼƬ����ɹ�. <br />"
		 s_Content = Replace(s_Content,RemoteFileurl,SaveFileName)
		 else 
		 Response.write ""&RemoteFileurl&" ͼƬ����<font color='#FF0000'>ʧ��</font>.<br />" 
		 end if
     Next
     ReplaceRemoteUrl = s_Content
End Function

function SaveRemoteFile(s_LocalFileName,s_RemoteFileUrl)
     Dim Ads, Retrieval, GetRemoteData
     On Error Resume Next
     Set Retrieval = Server.CreateObject("Microso" & "ft.XM" & "LHTTP")
     With Retrieval
         .Open "Get", s_RemoteFileUrl, False, "", ""
         .Send
         GetRemoteData = .ResponseBody
     End With
     Set Retrieval = Nothing
     Set Ads = Server.CreateObject("Ado" & "db.Str" & "eam") 
     With Ads
         .Type = 1
         .Open
         .Write GetRemoteData
         .SaveToFile Server.MapPath(s_LocalFileName), 2
         .Cancel()
         .Close()
     End With
     Set Ads=nothing	 
	 if err <> 0 then 
	 SaveRemoteFile = false 
	 err.clear 
	 else 
	 SaveRemoteFile = true 
	 end if
End Function
Function IsObjInstalled(s_ClassString)
     On Error Resume Next
     IsObjInstalled = False
     Err = 0
     Dim xTestObj
     Set xTestObj = Server.CreateObject(s_ClassString)
     If 0 = Err Then IsObjInstalled = True
     Set xTestObj = Nothing
     Err = 0
End Function

%>



