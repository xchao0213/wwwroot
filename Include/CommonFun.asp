<% 
'是否为已安装对象
Function isInstallObj(objname)
	dim isInstall,obj
	On Error Resume Next
	set obj=server.CreateObject(objname)
	if Err then 
		isInstallObj=false : err.clear 
	else 
		isInstallObj=true:set obj=nothing
	end if
End Function

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

Function ipvalue(txt,length)
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
txt = left(trim(txt),ii)&"***" 
exit for
end if
next
ipvalue= txt
else
ipvalue= ""
end if
End Function
'时间格试化
Function formatDate(Byval t,Byval ftype)	
  dim y, m, d, h, mi, s
  formatDate=""
  If IsDate(t)=False Then Exit Function
  y=cstr(year(t))
  m=cstr(month(t))
  If len(m)=1 Then m="0" & m
  d=cstr(day(t))
  If len(d)=1 Then d="0" & d
  h = cstr(hour(t))
  If len(h)=1 Then h="0" & h
  mi = cstr(minute(t))
  If len(mi)=1 Then mi="0" & mi
  s = cstr(second(t))
  If len(s)=1 Then s="0" & s
  select case cint(ftype)
  case 1		' yyyy-mm-dd
	  formatDate=y & "-" & m & "-" & d
  case 2			' yy-mm-dd	
	  formatDate=right(y,2) & "-" & m & "-" & d
  case 3		' mm-dd		
	  formatDate=m & "-" & d
  case 4		' yyyy-mm-dd hh:mm:ss
	  formatDate=y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  case 5		' hh:mm:ss
	  formatDate=h & ":" & mi & ":" & s
  case 6		' yyyy年mm月dd日
	  formatDate=y & "年" & m & "月" & d & "日"
  case 7		' yyyymmdd
	  formatDate=y & m & d
  case 8		'yyyymmddhhmmss
	  formatDate=y & m & d & h & mi & s
  case 9		' yyyy-mm
	  formatDate=y & "-" & m
  case 10		' dd
	  formatDate=d
  case 11		' mm
	  formatDate=m
  end select
End Function

'是否为空
Function isNul(str)
	if isnull(str) or str=""  then isNul=true else isNul=false
End Function

'是否为数字
Function isNum(str)
	if not isNul(str) then  isNum=isnumeric(str) else isNum=false
End Function

'获取扩展名
Function suffix(str)
	dim ext : str=trim(""&str) : ext=""
	if str<>"" then
		if instr(" "&str,"?")>0 then:str=mid(str,1,instr(str,"?")-1):end if
		if instrRev(str,".")>0 then:ext=mid(str,instrRev(str,".")):end if
	end if
	suffix=ext
End Function

'全角转换成半角
Function convertString(Str)
	Dim strChar,intAsc,strTmp,i
	For i = 1 To Len(Str)
      strChar = Mid(Str, i, 1)
      intAsc = Asc(strChar)
      If (intAsc>=-23648 And intAsc<=-23553) Then 
         strTmp = strTmp & Chr(intAsc+23680)
      Else
         strTmp = strTmp & strChar 
      End if    
    Next
	ConvertString=strTmp
End Function


'图片水印
'waterMarkImg(saveImgPath,waterMarkLocation)
Function waterMarkImg(saveImgPath,location)
dim sAllowMarkExt:sAllowMarkExt = ".jpg,.png,.gif,.jpeg,.bmp"
If InStr(sAllowMarkExt, Mid(saveImgPath, InStrRev(saveImgPath, "."), Len(saveImgPath))) = 0 Then Exit Function

	If Not isInstallObj("Persits.Jpeg") Then exit function
	set jpeg = Server.CreateObject("Persits.Jpeg")	
	strWidth=len(waterMarkFont)*13 : strHeight=3 	
	jpeg.Open Server.MapPath(saveImgPath)
	If  jpeg is nothing then exit function	
	if jpeg.width <200 and jpeg.height<200 then exit function
	'为图片加入水印功能
	jpeg.Canvas.Font.Color = &H000000 ' 颜色,这里是设置成:黑 
	jpeg.Canvas.Font.Family = "方正隶变简体"  ' 设置字体 
	jpeg.Canvas.Font.Bold = False '是否设置成粗体 
	jpeg.Canvas.Font.Size = 26 '字体大小 
	jpeg.Canvas.Font.Quality = 4 ' 文字清晰度		
	select case location
		case "1" : jpeg.Canvas.Print 5 , strHeight, waterMarkFont
		case "2" : jpeg.Canvas.Print (jpeg.width-strWidth) / 2, strHeight, waterMarkFont
		case "3" : jpeg.Canvas.Print jpeg.width-strWidth-5, strHeight, waterMarkFont
		case "4" : jpeg.Canvas.Print 5 , (jpeg.height-strHeight)/2, waterMarkFont
		case "5" : jpeg.Canvas.Print (jpeg.width-strWidth) / 2, (jpeg.height-strHeight)/2, waterMarkFont
		case "6" : jpeg.Canvas.Print jpeg.width-strWidth-5, (jpeg.height-strHeight)/2, waterMarkFont
		case "7" : jpeg.Canvas.Print 5 , jpeg.height-40, waterMarkFont
		case "8" : jpeg.Canvas.Print (jpeg.width-strWidth) / 2, jpeg.height-40, waterMarkFont
		case else : jpeg.Canvas.Print jpeg.width-strWidth-5, jpeg.height-40, waterMarkFont
	end select
	
	jpeg.Save Server.MapPath(saveImgPath)    ' 保存文件
	set jpeg=Nothing
End Function

Function imgurl(url)
if url<>"" then imgurl=dir&right(url,Len(url)-1) else imgurl="../images/nologo.jpg"
End Function
%>