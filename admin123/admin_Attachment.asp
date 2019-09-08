<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(32)
Response.Buffer=true 
On Error Resume Next
Server.ScriptTimeOut = 5000
Response.Expires=0 
Dim StartTime,IsReplace,ImageFolder
IsReplace = true	'是否过滤编辑时文件的<textarea></textarea>标记，如不过滤遇到有<textarea>标记的文件时编辑有显示不完全的现象
ImageFolder = "fso_image"     '存放本程序图标的文件夹，如果要更换图标文件夹请更新这里就可以拉
StatrTime = Timer()%>
<html>
<head>
<title>高屋置业附件在线管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
body{ margin-left:10px}
.fonts{font-size:9pt;line-height:25px}
.button{padding:2px;height:20px;background-color:#3A6592;border:1px #025582 solid;color: #FFFFFF;font-size:12px;font-family: "宋体";}
.TextBox {border-top-width: 1px;border-right-width: 1px;border-bottom-width: 1px;border-left-width: 1px;border-top-style: solid;border-right-style: solid;border-bottom-style: solid;border-left-style: solid;border-top-color: #666666;border-right-color: #CCCCCC;border-bottom-color: #CCCCCC;border-left-color: #666666;padding: 2px;height: 300px;font-size:12px;font-family: "宋体";}
.InputBox { border-top-width:1px;  border-left-width:1px; border-right-width:1px; border-bottom-width:1px;  border-top-color:#000000; border-left-color:#000000;  border-bottom-color:#000000;border-right-color:#000000; padding:2px;height:20px;}
Input, Select, TextArea {font-family: "宋体";font-size: 12px;text-decoration: none;}
img {border:0px}
a:link {color: #000000; text-decoration: none}
a:visited {color: #000000; text-decoration: none}
a:hover {color: #FF0000; text-decoration: underline}
</style>
<script language=javascript>
function Check()
{
  if(confirm("确定要保存文件么？\n此操作不可恢复！")){
   return true;
   }
  else{
   return false;
  }
}
</script></head>
<body>
<%
Sub IsErr()
	If Err = -2147221005 Then
		Response.write "这台服务器不支持FSO,故本程序无法运行"
		Response.end
	End If
End Sub
'*******************************************
'函数作用：取得文件的后缀名
'*******************************************
Function UpDir(ByVal D)
	Dim UDir
	If Len(D) = 0 then Exit Function
	UDir=Left(D,InStrRev(D,"\")-1)
	UpDir=UDir
End Function
'*******************************************
'函数作用：取得当前页的URL，
'		   为文件添加正确的链接
'*******************************************

Function FileUrl(url,D)
	Dim PageUrl,PUrl
	PageUrl="http://"& Request.ServerVariables("SERVER_NAME")&""
	PUrl="/uploadfile/"
	PageUrl=PageUrl & PUrl &Mid(D,2,Len(""&D&"")) & "/" & url
	FileUrl=PageUrl
End Function
'*******************************************
'函数作用：格式化文件的大小
'*******************************************
Function GetFileSize(size)
	Dim FileSize
	FileSize=size / 1024
	FileSize=FormatNumber(FileSize,2)
	If FileSize < 1024 and FileSize > 1 then
		GetFileSize="<font color=red>"& FileSize & "</font>&nbsp;KB"
	ElseIf FileSize >1024 then
		GetFileSize="<font color=red>"& FormatNumber(FileSize / 1024,2) & "</font>&nbsp;MB"
	Else
		GetFileSize="<font color=red>"& Size & "</font>&nbsp;Bytes"
	End If
End Function
'*******************************************
'函数作用：取得文件的后缀名
'*******************************************
Function GetExtensionName(name)
	Dim FileName
	FileName=Split(name,".")
	GetExtensionName=FileName(Ubound(FileName))
End Function
'*******************************************
'函数作用：返回文件类型
'*******************************************
Function GetFileIcon(name)
	Dim FileName,Icon
	FileName=Lcase(GetExtensionName(name))
	Select Case FileName
         Case "asp"
			 Icon = "asp.gif"
		 Case "bmp"
			 Icon = "bmp.gif"
		 Case "doc"
		 	Icon = "doc.gif"
		 Case "exe"
			 Icon = "exe.gif"
		 Case "gif"
			 Icon = "gif.gif"
		 Case "jpg"
			 Icon = "jpg.gif"
		 Case "chm"
			 Icon = "chm.gif"
		 Case "htm","html"
			 Icon = "htm.gif"
		 Case "log"
			 Icon = "log.gif"
		 Case "mdb"
		 	Icon = "mdb.gif"
		 Case "swf"
			 Icon = "swf.gif"
		 Case "txt"
		 	Icon = "txt.gif"
		 Case "wav"
			 Icon = "wav.gif"
		 Case "xls"
			 Icon = "xls.gif"
		 Case "rar","zip"
		 	Icon = "zip.gif"
		 Case "css"
		 	Icon = "css.gif"
		 Case Else
			Icon = "none.gif"
  End Select
  GetFileIcon=Icon
End Function
'*******************************************
'过程作用：删除选定的文件或文件夹
'*******************************************
Sub DelAll()
	Dim FolderId,FileId,ThisDir,FileNum,FolderNum,FilePath,FolderPath
	FolderId = Split(Request.Form("FolderId"),",")
	FileId = Split(Request.Form("FileId"),",")
	ThisDir = trim(Request.Form("ThisDir"))
	FileNum=0
	FolderNum=0
	If Ubound(FolderId) <> -1 then		'删除文件夹
		For i = 0 to Ubound(FolderId)
			FolderPath = Server.MapPath("/uploadfile/") & ThisDir & "\" & trim(FolderId(i))
			If obj.FolderExists(FolderPath) then
				obj.DeleteFolder FolderPath,true
				FolderNum = FolderNum + 1
			End If
		Next
	End If
	If Ubound(FileId) <> -1 then		'删除文件
		For j = 0 to Ubound(FileId)
			FilePath = Server.MapPath("/uploadfile/") & ThisDir & "\" & trim(FileId(j))
			If obj.FileExists(FilePath) then
				obj.DeleteFile FilePath,true
				FileNum = FileNum + 1
			End If
		Next
	End If
	Response.write "<script>alert('\n恭喜,删除成功\n\n"& FolderNum &" 个文件夹被删除\n"& FileNum &" 个文件被删除');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
End Sub
'*******************************************
'过程作用：使选定的文件或文件夹改名
'*******************************************
Sub Rname()
Dim ThisDir,FolderName,NewName,OldN
ThisDir = Trim(Request.Form("ThisDir"))
FolderName = Trim(Request.Form("FolderId"))
FileName = Trim(Request.Form("FileId"))
NewName = Trim(Request.QueryString("NewName"))

If len(FileName) <> 0 then	'文件改名
	Newn = Server.MapPath("/uploadfile/") & ThisDir & "\" & NewName
	OldN = Server.MapPath("/uploadfile/") & ThisDir & "\" & FileName
	If not obj.FileExists(Newn) then
	obj.MoveFile OldN,Newn
	Response.write "<script>alert('【"&FileName&"】文件已经成功改名【"&NewName&"】');window.location.href=('"&Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
Else
	Response.write "<script>alert('有同名文件，请换个文件名');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
End If
End If	
End Sub
'*******************************************
'过程作用：新建文件
'*******************************************
Sub NewFile()
	Dim NewFile,NewFilePath
	NewFilePath = Trim(Request.Form("ThisDir"))
	NewFile = Trim(Request.Form("NewFileName"))
	NewFilePath = Server.MapPath("/uploadfile/") & NewFilePath & "\" & NewFile
	If not obj.FileExists(NewFilePath) and not obj.FolderExists(NewFilePath) then
		Set objFi = obj.CreateTextFile(NewFilePath)
			objFi.Writeline
		objFi.close
		Set objFi = nothing
		Response.write "<script>alert('建立文件成功');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	Else
		Response.write "<script>alert('有同名文件，请换个文件名');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	End if
End Sub
'*******************************************
'过程作用：新建文件夹
'*******************************************
Sub NewFolder()
	Dim NewFolder,NewFolderPath
	NewFolderPath = Trim(Request.Form("ThisDir"))
	NewFolder = Trim(Request.Form("NewFolderName"))
	NewFolderPath = Server.MapPath("/uploadfile/") & NewFolderPath & "\" & NewFolder
	If not obj.FolderExists(NewFolderPath) then
		obj.CreateFolder(NewFolderPath)
		Response.write "<script>alert('建立文件夹成功');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	Else
		Response.write "<script>alert('有同名文件夹，请换个文件夹名');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	End if
End Sub

Sub Css()
%>

<%
End Sub
'*******************************************
'过程作用：编辑文件
'*******************************************
Sub Edit()
	Dim FilePath,FileName,action
	Set obj= Server.CreateObject("Scripting.FileSystemObject")
	IsErr
	action=Trim(Request.QueryString("action"))
	If action = ("Save") then	'保存文件
		Dim FileSave
		FilePath = trim(Request.QueryString("FilePath"))
		FileAll = trim(Request.Form("FileAll"))
		If IsReplace then FileAll = Replace(FileAll,"\\textarea\\","textarea")
		If obj.FileExists(FilePath) then
			Set FileSave = obj.OpenTextFile(FilePath,2)
				FileSave.Write(FileAll)
			FileSave.Close
			Response.write "<script>if(confirm('文件已经保存，是否关闭本页')){window.close();}else{history.back();}</script>"
		Else
			Response.write "<script>alert('发生错误，文件已经被删除或者损坏！');window.close()</script>"
		End If
	ElseIf action = ("Edit") then	'读取文件
		Dim FileAll
		FilePath = Trim(Request.Form("ThisDir"))
		FileName = Trim(Request.Form("FileId"))
		FilePath1 = Server.MapPath("/uploadfile/") & FilePath & "\" & FileName
		If obj.FileExists(FilePath1) then
			Set FileOpen = obj.OpenTextFile (FilePath1,1)
				FileAll = FileOpen.ReadAll
			FileOpen.close
		If IsReplace then FileAll = Replace(FileAll,"textarea","\\textarea\\")
		Else
			Response.write "<script>alert('发生错误，文件已经被删除或者损坏！');window.close()</script>"
		End If
%>

<% Call Css %>
<form name="form" method="post" action="?action=Save&FilePath=<%=Server.MapPath("/uploadfile/") & FilePath & "\" & Filename %>" onSubMit="return Check()">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="fonts">
 <tr> 
  <td>系统主目录：<a href=?action=Open&Dir= title=返回到系统主目录><font color=FF6600><b><%=Server.MapPath("/uploadfile/")%></b></font></a></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr> 
   <td height="20" valign="bottom"><a href="?action=Open&Dir="><input type="button" value="返回到主目录" class=button style="cursor:hand"></a>&nbsp;&nbsp;当前目录：<%=Server.MapPath("/uploadfile/") & FilePath %></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr>
   <td height="30" valign="middle">&nbsp;文件名:&nbsp;<font color=red><b><%=FileName%></b></font>&nbsp;
	<select name="select" onChange="FileAll.style.fontSize=this.options[this.options.selectedIndex].value">
	 <option selected value="12px">字体大小</option>
	 <option value="12px">12px</option>
	 <option value="14px">14px</option>
	 <option value="16px">16px</option>
	</select>
	<select name="select2" onChange="FileAll.style.color=this.options[selectedIndex].value">
	 <option selected value="#000000">颜色</option>
	 <option value="#666666">灰色</option>
	 <option value="#ff0000">红色</option>
	 <option value="#087100">绿色</option>
	</select>
	&nbsp;&nbsp;&nbsp;&nbsp; </td>
 </tr>
 <tr>
  <td bgcolor="B7CECD" height="1" id="a11"></td>
 </tr>
 <tr>
  <td><table cellpadding=3 cellspac=0 class=fonts><tr><td>
	   <textarea id="FileAll" name="FileAll" cols="100" rows="20" class=textbox style="word-break: break-all; width: 700px; height: 380px;"><%=FileAll%></textarea>
</td><tr><td height=0></td></tr><tr><td>
	   <input type="submit" name="Submit" value=" 保存文件 " class=button>
	   <input type="reset" name="Submit2" value=" 撤消修改 " class=button>
	<input type="button" name="Submit3" value=" 关闭窗口 " class=button onClick="window.close()">
	</td></tr></table>
   </td>
 </tr>
</table>
   </form>

<%	End If
	Set obj = nothing
	End Sub
'****************************************
'函数定义部分结束
'****************************************
Dim obj,objFi,FileType,FileSize,FileTime,Path
	action=Trim(Request.QueryString("action"))
	Set obj=Server.CreateObject("Scripting.FileSystemObject")
	IsErr
	If action = "Del" then
		Call DelAll
	ElseIf action = "NewFile" then
		Call NewFile
	ElseIf action = "NewFolder" then
		Call NewFolder
	ElseIf action = "Rname" then
		Call Rname
	ElseIf action = "Edit" then
		Call Edit
	ElseIf action = "Save" then
		Call Edit
	Else
		Dir=Trim(Request.QueryString("Dir"))
		Path = Server.MapPath("/uploadfile/") & Dir
		Set objFi = obj.GetFolder(Server.MapPath("/uploadfile/"))
			objFiSize = objFi.size	'空间大小统计
		Set	objFi = nothing
		Set	objFi = obj.GetFolder(Path)
%>

<% Call Css %>
<script language=javascript>
function Checked()
{
	var j = 0
	for(i=0;i < document.form.elements.length;i++){
		if(document.form.elements[i].name == "FileId" || document.form.elements[i].name == "FolderId"){
			if(document.form.elements[i].checked){
				j++;
			}
		}
	}
	return j;
}
function CheckAll1()
{
	for(i=0;i<document.form.elements.length;i++)
	{
		if(document.form.elements[i].checked){
			document.form.elements[i].checked=false;
			document.form.CheckAll.checked=false;
		}
		else{
			document.form.elements[i].checked = true;
			document.form.CheckAll.checked = true;
		}
	}
}
function DelAll()
{
	if(Checked()  <= 0){
		alert("您必须选择其中的一个文件或文件夹");
	}	
	else{
		if(confirm("确定要删除选择的文件或文件夹么？\n此操作不可以恢复！")){
			form.action="?action=Del";
			form.submit();
		}
	}
}
function Edit()
{
	if(Checked() == 0){
		alert("您必须选择其中的一个文件");
	}
	else{
		if(Checked() != 1){
			alert("只能选择一个文件（文本文件）");
		}
		else{
			for(i=0;i < document.form.elements.length;i++){
			if(document.form.elements[i].name == "FileId" && document.form.elements[i].checked){
				form.action="?action=Edit";
				form.target="self";
				form.submit();
				break;
			}
			else if(document.form.elements[i].name == "FolderId" && document.form.elements[i].checked){
				alert("不能编辑文件夹")
				break;
			}
			}
		}
	}
}
function Rname()
{
	if(Checked() == 0){
		alert("您必须选择一个文件或文件夹");
	}
	else{
		if(Checked() != 1){
			alert("只能选择一个文件或一个文件夹");
		}
		else{
			for(i=0;i < document.form.elements.length;i++){
				if(document.form.elements[i].name == "FolderId" && document.form.elements[i].checked){
					var j = prompt("请输入新文件夹名",document.form.elements[i].value)
					break;
				}
				else if(document.form.elements[i].name == "FileId" && document.form.elements[i].checked){
					var j = prompt("请输入新文件名",document.form.elements[i].value)
					break;
				}
			}
			if(j != "" && j != null){
				if(IsStr(j) == j.length){
					form.action="?action=Rname&NewName=" + j;
                                                      form.target="_self";
					form.submit();
				}
				else{
					alert("新名称不符合标准，只能是字母、数字、点和下划线的组合,\n不能含有汉字、空格和其他符号");
				}
			}
		}
	}
}
function IsStr(w)
{
	var str = "abcdefghijklmnopqrstuvwxyz_1234567890."
	 w = w.toLowerCase();
	var j = 0;
	for(i=0;i < w.length;i++){
		if(str.indexOf(w.substr(i,1)) != -1){
			j++;
		}
	}
	return j;
}
function NewFile(form,i)
{
	if(i == 1){
		if(form.NewFolderName.value == ""){
			alert("文件夹名不能为空");
		}
		else{
			if(IsStr(form.NewFolderName.value) == form.NewFolderName.value.length){
				form.action="?action=NewFolder";
				form.submit();
			}
			else{
				alert("文件夹名不符合标准，只能是字母、数字、点和下划线的组合,\n不能含有汉字、空格和其他符号");
			}
		}
	}
	else{
		if(form.NewFileName.value == ""){
			alert("文件名不能为空");
		}
		else{
			if(IsStr(form.NewFileName.value) == form.NewFileName.value.length){
				form.action="?action=NewFile";
				form.submit();
			}
			else{
				alert("文件名不符合标准，只能是字母、数字、点和下划线的组合,\n不能含有汉字、空格和其他符号");
			}
		}
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">上传附件管理</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" class="fonts">

 <tr> 
  <td style="font-size:14px">系统主目录：<a href=?action=Open&Dir= title=返回到系统主目录><font color=FF6600><b><%=Server.MapPath("/uploadfile/")%></b></font></a>&nbsp;&nbsp;&nbsp;空间占用：<%=GetFileSize(objFiSize)%></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr> 
  <td height="20" valign=bottom><a href=?action=Open&Dir=<%=UpDir(Dir)%>><input type="button" value="返回到上一目录" class=button style="cursor:hand"></a>&nbsp;
        &nbsp;当前目录：<%=Server.MapPath("/admin/") & Dir %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;占用空间：<%=GetFileSize(objFi.size)%>&nbsp;&nbsp;其中包含&nbsp;<font color=red><%=objFi.SubFolders.count%></font>&nbsp;个文件夹；&nbsp;<font color=red><%=objFi.Files.count%></font>&nbsp;个文件</td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr><%strHtml=Trim(Request.QueryString("action"))
 if strHtml="admin" then%>
 <form name="form1" method="post">
  <tr> 
   <td height="60" valign="middle">新建文件夹： 
	<input type="text" name="NewFolderName" size="15" class=InputBox  maxlength="50">
	&nbsp; 
	<input type="button" name="Submit4" value="新建文件夹" class=button style="cursor:hand" onClick="NewFile(this.form,1)">
	<font color="990033"> 
	<input type="hidden" name="ThisDir" value="<%=Dir%>">
	</font>新建文件：&nbsp;&nbsp; 
	<input type="text" name="NewFileName" size="15" class=InputBox  maxlength="50">
	&nbsp; 
	<input type="button" name="Submit5" value=" 新建文件 " class=button style="cursor:hand" onClick="NewFile(this.form,2)">
   </td>
  </tr>
 </form><% End If %>
 <tr> 
  <td valign="top"> 
   <table width="99%" border="0" cellspacing="1" cellpadding="0" bgcolor="B4E2EF" class="fonts">
	<form name="form" method="post" >
	 <tr bgcolor="F4F4F4"> 
	  <Td width="6%" align="center">&nbsp;</td>
	  <td width="39%"><font color="990033">&nbsp;文件/文件夹名 </font></td>
	  <td width="13%" align="center"><font color="990033">文件浏览</font></td>
	  <td width="13%" align="center"><font color="990033">类型</font></td>
	  <td width="15%" align="center"><font color="990033">文件大小</font></td>
	  <td width="27%" align="center"><font color="990033">最后修改时间</font></td>
	 </tr>
<%	For Each DirFolder in objFi.SubFolders
FolderName=DirFolder.name
FolderSize=GetFileSize(DirFolder.size)
FolderTime=DirFolder.DateLastModified
%>
	 <tr bgcolor="#FFFFFF"> 
	  <td width="6%" align="center"> 
	   <input type="checkbox" name="FolderId" value="<%=FolderName%>">	  </td>
	  <td width="39%">&nbsp;<a href=?action=Open&Dir=<%=Dir%>\<%=FolderName%>><%=FolderName%></a></td>
	  <td width="15%" align="center">无</td>
	  <td width="13%" align="center"><img src="<%=ImageFolder%>/ClosedFolder.gif" width="16" height="16" alt="文件夹"></td>
	  <td width="15%" align="center"><%=FolderSize%></td>
	  <td width="27%" align="center"><%=FolderTime%></td>
	 </tr>
	 <%	Next %>
	 <%	For Each DirFiles in objFi.Files
		FileName=DirFiles.name
		FileType=GetFileIcon(FileName)
		FileSize=GetFileSize(DirFiles.size)
		FileTime=DirFiles.DateLastModified%>
	 <tr bgcolor="#FFFFFF"> 
	  <td width="6%" align="center"> 
	   <input type="checkbox" name="FileId" value="<%=FileName%>">	  </td>
	  <td width="39%">&nbsp;<a href=<%=FileUrl(FileName,Dir)%> target=_blank><%=FileName%></a></td>
	   <td width="15%" align="center" style="padding:2px 0px"><a href=<%=FileUrl(FileName,Dir)%> target=_blank>
	   <%
		la=split(FileName,"/")
		num=ubound(la)
		lb=split(la(num),".")
		num2=ubound(lb)
		suffix="."&lb(num2)&""
		if suffix=".swf" then%>
		SWF文件请点击预览 
		<%elseif suffix=".asp" then%>
		<%else%>
		 <img src="<%=FileUrl(FileName,Dir)%>" alt="点击查看文件大图" width="115" height="50">
		<%end if%>	   
	   </a></td>
	  <td width="13%" align="center"><Img src=<%=ImageFolder%>/<%=FileType%> alt="文件"></td>
	  <td width="15%" align="center"><%=FileSize%></td>
	  <td width="27%" align="center"><%=FileTime%></td>
	 </tr>
	 <%	Next %>
	 <tr bgcolor="#FFFFFF"> 
	  <td width="6%" align="center"> 
	   <input type="checkbox" name="CheckAll" value="checkbox" onClick="CheckAll1()" title=全部选择 style="cursor:hand">	  </td>
	  <td colspan="5" height="30">
<%strHtml=Trim(Request.QueryString("action"))
 if strHtml="admin" then%>
	   <input type="button" name="Submit" value="编 辑" class=button style="cursor:hand" onClick="Edit()"  title=编辑>
	   <input type="button" name="Submit3" value="重命名" class=button style="cursor:hand" onClick="Rname()"  title=重命名>
	   <% End If %>
	   <input type="button" name="Submit2" value="删 除" class=button style="cursor:hand" onClick="DelAll()"  title=删除>
	   <font color="990033"> <input type="hidden" name="ThisDir" value="<%=Dir%>"></font>
	   <input name="button" type="button" class=button style="cursor:hand" onClick="window.location.href='admin_file.asp'"  value="清理沉余文件"></td>
	 </tr>
	</form>
   </table>
  </td>
 </tr>
</table>
<br>
</td>
  </tr>
</table>
</body>
</html>
<%End If
Set objFi = nothing
Set obj = nothing
%>