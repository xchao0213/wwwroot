<%
'HTML 编码
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

'HTML 解码
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


'=========格式化日期==================
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
		  DateString = Month(DateStr)&"月"&Day(DateStr)&"日"
	  Case "12"
		  DateString = Day(DateStr)&"日"&Hour(DateStr)&"时"
	  case "13"
		  DateString = Day(DateStr)&"日"&Hour(DateStr)&"点"
	  Case "14"
		  DateString = Hour(DateStr)&"时"&Minute(DateStr)&"分"
	  Case "15"
		  DateString = Hour(DateStr)&":"&Minute(DateStr)
	  Case "16"
		  DateString = Year(DateStr)&"年"&Month(DateStr)&"月"&Day(DateStr)&"日"
	  Case Else
	  	  DateString = DateStr
	 End Select
	 DateFormat = DateString
 End Function


'去除html标记函数内容如下:
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
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串带...
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
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function


'*************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串不带...
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

'读取文件
Function ReadtxtFiles(Filename)
   Dim fso, f1, ts, s
   Const ForReading = 1
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set ts = fso.OpenTextFile(Filename, ForReading)
   ReadtxtFiles = ts.ReadAll
   ts.Close
End Function

'创建文件
Sub CreateFile(Filename,temp)
 '  Dim fso, tf ,Filename ,temp
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set tf = fso.CreateTextFile(Filename, True)
   ' 写一行，并且带有新行字符。
'   tf.WriteLine("Testing 1, 2, 3.") 
   '向文件写三个新行字符。        
'   tf.WriteBlankLines(3) 
   '写一行。
'   tf.Write ("This is a test.") 
   tf.write temp
   tf.Close
End Sub

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'pos=InStr(1,"abcdefg","cd") 
'则pos会返回3表示查找到并且位置为第三个字符开始。
'这就是“查找”的实现，而“查找下一个”功能的
'实现就是把当前位置作为起始位置继续查找。
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
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
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
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub

'********************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
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
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
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
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
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
'函数名：SendMail
'作  用：用Jmail组件发送邮件
'参  数：ServerAddress  ----服务器地址
'        AddRecipient  ----收信人地址
'        Subject       ----主题
'        Body          ----信件内容
'        Sender        ----发信人地址
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>没有安装JMail组件</li>"
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
'函数名：ReplaceRemoteUrl
'作  用：远程下载图片
'****************************************************

Const sFileExt="jpg|gif|bmp|png"
Function ReplaceRemoteUrl(sHTML, sSaveFilePath, sFileExt)
	 If Not fso.FolderExists(Server.mappath(uploadpath)) Then
	  fso.CreateFolder(Server.mappath(uploadpath))'创建目录
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
		 'response.Write ""&SaveFileName&" 图片保存成功. <br />"
		 s_Content = Replace(s_Content,RemoteFileurl,SaveFileName)
		 else 
		 Response.write ""&RemoteFileurl&" 图片保存<font color='#FF0000'>失败</font>.<br />" 
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



