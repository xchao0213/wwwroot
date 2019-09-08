<!--#include file="../Include/conn.asp" -->
<!--#include file="seeion.asp" -->
<%call chkAdmin(33)
SiteUp="uploadfile/image/"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equtv="Content-Type" content="text/html; charset=gb2312" />
<title>网站后台管理--文件清理</title>
<link href="images/style.css" type=text/css rel=stylesheet>
</head>
<body>
<%	
if request("action")="del" then
call DelOver()
else
call Main()
end if
Sub Main%>
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#666666">
    <tr> 
      <td height="27" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">上传文件清理</div></td>
    </tr><tr>
<td height=30 bgcolor="#FFFFFF" class="td" style="padding-left:10px"><b>注意：</b>请谨慎使用本功能，操作前最好先备份下上传文件夹，以免文件丢失，清理是清理掉不在数据库中存在或正在使用上传的文件。<br>
<br>
如果你的文章特别多(数万)，清理起来可能会非常慢，不推荐使用本功能。
<br><br>当前上传文件夹为：<%=SiteUp%> <input type="button" value="开始清理" class="btn" onClick="if(confirm('确定要清除文件夹<%=SiteUp%>中的垃圾文件?\n\n请谨慎操作！'))location.href='?action=del';return false;" />
</td>
</tr>
</table>
<%
End Sub

Sub DelOver
	Dim Sql,Rs,strFiles,ItemIntro
	strFiles="|"&dir&SiteUp&"|"&SiteLogo
	ItemIntro=""
	
	Sql = "Select body,img From news"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End if
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Sql = "Select body From about"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img From project"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img From project"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Sql = "Select body,url From download"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Sql = "Select img From ad"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img From flash"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select logo From link"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img From indexflash"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
		
	Sql = "Select img,body From Products"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img,body From video"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img,body From tuandui"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select ImagePath,img,body From [case]"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
		IF Rs(2) <> "" Then
			strFiles=strFiles&"|"&Rs(2)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Sql = "Select img,ImagePath,body From photo"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
		IF Rs(2) <> "" Then
			strFiles=strFiles&"|"&Rs(2)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
		Sql = "Select img,body From video"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	Dim tempStr, tempi, TempArray,UpFileType,regEx,Matches,Match
		UpFileType = "gif|jpg|jpeg|bmp|png"
		Set regEx=New Regexp
		regEx.Ignorecase=True
		regEx.Global=True
        regEx.Pattern = "<img.+?[^\>]>" '查询内容中所有 <img..>
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '累计数组
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '分割数组
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "src\s*=\s*.+?\.(" & UpFileType & ")" '查询src =内的链接
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '累加得到 链接加$Array$ 字符
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "src\s*=\s*" '过滤 src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		Set regEx=Nothing
        strFiles = strFiles & tempStr
	    strFiles = LCase(strFiles)
		Dim i,theFolder,fso,theFile,theSubFolder
		Set Fso=CreateObject("Scripting.FileSystemObject")
		 i = 0
     

    Set theFolder = fso.GetFolder(Server.MapPath(dir&SiteUp))
    For Each theFile In theFolder.Files
        If InStr(strFiles, LCase(theFile.name)) <= 0 Then
            theFile.Delete True
            i = i + 1
        End If
    Next
    For Each theSubFolder In theFolder.SubFolders
        For Each theFile In theSubFolder.Files
            If InStr(strFiles, LCase(theSubFolder.name & "/" & theFile.name)) <= 0 Then
                theFile.Delete True
                i = i + 1
            End If
        Next
    Next

	Response.Write "<script>alert('本次共清理"&I&"个垃圾文件！');history.go(-1);</script>" 

End Sub
%>
</body>
</html>