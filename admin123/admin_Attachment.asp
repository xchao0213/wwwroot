<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(32)
Response.Buffer=true 
On Error Resume Next
Server.ScriptTimeOut = 5000
Response.Expires=0 
Dim StartTime,IsReplace,ImageFolder
IsReplace = true	'�Ƿ���˱༭ʱ�ļ���<textarea></textarea>��ǣ��粻����������<textarea>��ǵ��ļ�ʱ�༭����ʾ����ȫ������
ImageFolder = "fso_image"     '��ű�����ͼ����ļ��У����Ҫ����ͼ���ļ������������Ϳ�����
StatrTime = Timer()%>
<html>
<head>
<title>������ҵ�������߹���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
body{ margin-left:10px}
.fonts{font-size:9pt;line-height:25px}
.button{padding:2px;height:20px;background-color:#3A6592;border:1px #025582 solid;color: #FFFFFF;font-size:12px;font-family: "����";}
.TextBox {border-top-width: 1px;border-right-width: 1px;border-bottom-width: 1px;border-left-width: 1px;border-top-style: solid;border-right-style: solid;border-bottom-style: solid;border-left-style: solid;border-top-color: #666666;border-right-color: #CCCCCC;border-bottom-color: #CCCCCC;border-left-color: #666666;padding: 2px;height: 300px;font-size:12px;font-family: "����";}
.InputBox { border-top-width:1px;  border-left-width:1px; border-right-width:1px; border-bottom-width:1px;  border-top-color:#000000; border-left-color:#000000;  border-bottom-color:#000000;border-right-color:#000000; padding:2px;height:20px;}
Input, Select, TextArea {font-family: "����";font-size: 12px;text-decoration: none;}
img {border:0px}
a:link {color: #000000; text-decoration: none}
a:visited {color: #000000; text-decoration: none}
a:hover {color: #FF0000; text-decoration: underline}
</style>
<script language=javascript>
function Check()
{
  if(confirm("ȷ��Ҫ�����ļ�ô��\n�˲������ɻָ���")){
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
		Response.write "��̨��������֧��FSO,�ʱ������޷�����"
		Response.end
	End If
End Sub
'*******************************************
'�������ã�ȡ���ļ��ĺ�׺��
'*******************************************
Function UpDir(ByVal D)
	Dim UDir
	If Len(D) = 0 then Exit Function
	UDir=Left(D,InStrRev(D,"\")-1)
	UpDir=UDir
End Function
'*******************************************
'�������ã�ȡ�õ�ǰҳ��URL��
'		   Ϊ�ļ������ȷ������
'*******************************************

Function FileUrl(url,D)
	Dim PageUrl,PUrl
	PageUrl="http://"& Request.ServerVariables("SERVER_NAME")&""
	PUrl="/uploadfile/"
	PageUrl=PageUrl & PUrl &Mid(D,2,Len(""&D&"")) & "/" & url
	FileUrl=PageUrl
End Function
'*******************************************
'�������ã���ʽ���ļ��Ĵ�С
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
'�������ã�ȡ���ļ��ĺ�׺��
'*******************************************
Function GetExtensionName(name)
	Dim FileName
	FileName=Split(name,".")
	GetExtensionName=FileName(Ubound(FileName))
End Function
'*******************************************
'�������ã������ļ�����
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
'�������ã�ɾ��ѡ�����ļ����ļ���
'*******************************************
Sub DelAll()
	Dim FolderId,FileId,ThisDir,FileNum,FolderNum,FilePath,FolderPath
	FolderId = Split(Request.Form("FolderId"),",")
	FileId = Split(Request.Form("FileId"),",")
	ThisDir = trim(Request.Form("ThisDir"))
	FileNum=0
	FolderNum=0
	If Ubound(FolderId) <> -1 then		'ɾ���ļ���
		For i = 0 to Ubound(FolderId)
			FolderPath = Server.MapPath("/uploadfile/") & ThisDir & "\" & trim(FolderId(i))
			If obj.FolderExists(FolderPath) then
				obj.DeleteFolder FolderPath,true
				FolderNum = FolderNum + 1
			End If
		Next
	End If
	If Ubound(FileId) <> -1 then		'ɾ���ļ�
		For j = 0 to Ubound(FileId)
			FilePath = Server.MapPath("/uploadfile/") & ThisDir & "\" & trim(FileId(j))
			If obj.FileExists(FilePath) then
				obj.DeleteFile FilePath,true
				FileNum = FileNum + 1
			End If
		Next
	End If
	Response.write "<script>alert('\n��ϲ,ɾ���ɹ�\n\n"& FolderNum &" ���ļ��б�ɾ��\n"& FileNum &" ���ļ���ɾ��');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
End Sub
'*******************************************
'�������ã�ʹѡ�����ļ����ļ��и���
'*******************************************
Sub Rname()
Dim ThisDir,FolderName,NewName,OldN
ThisDir = Trim(Request.Form("ThisDir"))
FolderName = Trim(Request.Form("FolderId"))
FileName = Trim(Request.Form("FileId"))
NewName = Trim(Request.QueryString("NewName"))

If len(FileName) <> 0 then	'�ļ�����
	Newn = Server.MapPath("/uploadfile/") & ThisDir & "\" & NewName
	OldN = Server.MapPath("/uploadfile/") & ThisDir & "\" & FileName
	If not obj.FileExists(Newn) then
	obj.MoveFile OldN,Newn
	Response.write "<script>alert('��"&FileName&"���ļ��Ѿ��ɹ�������"&NewName&"��');window.location.href=('"&Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
Else
	Response.write "<script>alert('��ͬ���ļ����뻻���ļ���');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
End If
End If	
End Sub
'*******************************************
'�������ã��½��ļ�
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
		Response.write "<script>alert('�����ļ��ɹ�');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	Else
		Response.write "<script>alert('��ͬ���ļ����뻻���ļ���');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	End if
End Sub
'*******************************************
'�������ã��½��ļ���
'*******************************************
Sub NewFolder()
	Dim NewFolder,NewFolderPath
	NewFolderPath = Trim(Request.Form("ThisDir"))
	NewFolder = Trim(Request.Form("NewFolderName"))
	NewFolderPath = Server.MapPath("/uploadfile/") & NewFolderPath & "\" & NewFolder
	If not obj.FolderExists(NewFolderPath) then
		obj.CreateFolder(NewFolderPath)
		Response.write "<script>alert('�����ļ��гɹ�');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	Else
		Response.write "<script>alert('��ͬ���ļ��У��뻻���ļ�����');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
	End if
End Sub

Sub Css()
%>

<%
End Sub
'*******************************************
'�������ã��༭�ļ�
'*******************************************
Sub Edit()
	Dim FilePath,FileName,action
	Set obj= Server.CreateObject("Scripting.FileSystemObject")
	IsErr
	action=Trim(Request.QueryString("action"))
	If action = ("Save") then	'�����ļ�
		Dim FileSave
		FilePath = trim(Request.QueryString("FilePath"))
		FileAll = trim(Request.Form("FileAll"))
		If IsReplace then FileAll = Replace(FileAll,"\\textarea\\","textarea")
		If obj.FileExists(FilePath) then
			Set FileSave = obj.OpenTextFile(FilePath,2)
				FileSave.Write(FileAll)
			FileSave.Close
			Response.write "<script>if(confirm('�ļ��Ѿ����棬�Ƿ�رձ�ҳ')){window.close();}else{history.back();}</script>"
		Else
			Response.write "<script>alert('���������ļ��Ѿ���ɾ�������𻵣�');window.close()</script>"
		End If
	ElseIf action = ("Edit") then	'��ȡ�ļ�
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
			Response.write "<script>alert('���������ļ��Ѿ���ɾ�������𻵣�');window.close()</script>"
		End If
%>

<% Call Css %>
<form name="form" method="post" action="?action=Save&FilePath=<%=Server.MapPath("/uploadfile/") & FilePath & "\" & Filename %>" onSubMit="return Check()">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="fonts">
 <tr> 
  <td>ϵͳ��Ŀ¼��<a href=?action=Open&Dir= title=���ص�ϵͳ��Ŀ¼><font color=FF6600><b><%=Server.MapPath("/uploadfile/")%></b></font></a></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr> 
   <td height="20" valign="bottom"><a href="?action=Open&Dir="><input type="button" value="���ص���Ŀ¼" class=button style="cursor:hand"></a>&nbsp;&nbsp;��ǰĿ¼��<%=Server.MapPath("/uploadfile/") & FilePath %></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr>
   <td height="30" valign="middle">&nbsp;�ļ���:&nbsp;<font color=red><b><%=FileName%></b></font>&nbsp;
	<select name="select" onChange="FileAll.style.fontSize=this.options[this.options.selectedIndex].value">
	 <option selected value="12px">�����С</option>
	 <option value="12px">12px</option>
	 <option value="14px">14px</option>
	 <option value="16px">16px</option>
	</select>
	<select name="select2" onChange="FileAll.style.color=this.options[selectedIndex].value">
	 <option selected value="#000000">��ɫ</option>
	 <option value="#666666">��ɫ</option>
	 <option value="#ff0000">��ɫ</option>
	 <option value="#087100">��ɫ</option>
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
	   <input type="submit" name="Submit" value=" �����ļ� " class=button>
	   <input type="reset" name="Submit2" value=" �����޸� " class=button>
	<input type="button" name="Submit3" value=" �رմ��� " class=button onClick="window.close()">
	</td></tr></table>
   </td>
 </tr>
</table>
   </form>

<%	End If
	Set obj = nothing
	End Sub
'****************************************
'�������岿�ֽ���
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
			objFiSize = objFi.size	'�ռ��Сͳ��
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
		alert("������ѡ�����е�һ���ļ����ļ���");
	}	
	else{
		if(confirm("ȷ��Ҫɾ��ѡ����ļ����ļ���ô��\n�˲��������Իָ���")){
			form.action="?action=Del";
			form.submit();
		}
	}
}
function Edit()
{
	if(Checked() == 0){
		alert("������ѡ�����е�һ���ļ�");
	}
	else{
		if(Checked() != 1){
			alert("ֻ��ѡ��һ���ļ����ı��ļ���");
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
				alert("���ܱ༭�ļ���")
				break;
			}
			}
		}
	}
}
function Rname()
{
	if(Checked() == 0){
		alert("������ѡ��һ���ļ����ļ���");
	}
	else{
		if(Checked() != 1){
			alert("ֻ��ѡ��һ���ļ���һ���ļ���");
		}
		else{
			for(i=0;i < document.form.elements.length;i++){
				if(document.form.elements[i].name == "FolderId" && document.form.elements[i].checked){
					var j = prompt("���������ļ�����",document.form.elements[i].value)
					break;
				}
				else if(document.form.elements[i].name == "FileId" && document.form.elements[i].checked){
					var j = prompt("���������ļ���",document.form.elements[i].value)
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
					alert("�����Ʋ����ϱ�׼��ֻ������ĸ�����֡�����»��ߵ����,\n���ܺ��к��֡��ո����������");
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
			alert("�ļ���������Ϊ��");
		}
		else{
			if(IsStr(form.NewFolderName.value) == form.NewFolderName.value.length){
				form.action="?action=NewFolder";
				form.submit();
			}
			else{
				alert("�ļ����������ϱ�׼��ֻ������ĸ�����֡�����»��ߵ����,\n���ܺ��к��֡��ո����������");
			}
		}
	}
	else{
		if(form.NewFileName.value == ""){
			alert("�ļ�������Ϊ��");
		}
		else{
			if(IsStr(form.NewFileName.value) == form.NewFileName.value.length){
				form.action="?action=NewFile";
				form.submit();
			}
			else{
				alert("�ļ��������ϱ�׼��ֻ������ĸ�����֡�����»��ߵ����,\n���ܺ��к��֡��ո����������");
			}
		}
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">�ϴ���������</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" class="fonts">

 <tr> 
  <td style="font-size:14px">ϵͳ��Ŀ¼��<a href=?action=Open&Dir= title=���ص�ϵͳ��Ŀ¼><font color=FF6600><b><%=Server.MapPath("/uploadfile/")%></b></font></a>&nbsp;&nbsp;&nbsp;�ռ�ռ�ã�<%=GetFileSize(objFiSize)%></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr> 
  <td height="20" valign=bottom><a href=?action=Open&Dir=<%=UpDir(Dir)%>><input type="button" value="���ص���һĿ¼" class=button style="cursor:hand"></a>&nbsp;
        &nbsp;��ǰĿ¼��<%=Server.MapPath("/admin/") & Dir %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ռ�ÿռ䣺<%=GetFileSize(objFi.size)%>&nbsp;&nbsp;���а���&nbsp;<font color=red><%=objFi.SubFolders.count%></font>&nbsp;���ļ��У�&nbsp;<font color=red><%=objFi.Files.count%></font>&nbsp;���ļ�</td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr><%strHtml=Trim(Request.QueryString("action"))
 if strHtml="admin" then%>
 <form name="form1" method="post">
  <tr> 
   <td height="60" valign="middle">�½��ļ��У� 
	<input type="text" name="NewFolderName" size="15" class=InputBox  maxlength="50">
	&nbsp; 
	<input type="button" name="Submit4" value="�½��ļ���" class=button style="cursor:hand" onClick="NewFile(this.form,1)">
	<font color="990033"> 
	<input type="hidden" name="ThisDir" value="<%=Dir%>">
	</font>�½��ļ���&nbsp;&nbsp; 
	<input type="text" name="NewFileName" size="15" class=InputBox  maxlength="50">
	&nbsp; 
	<input type="button" name="Submit5" value=" �½��ļ� " class=button style="cursor:hand" onClick="NewFile(this.form,2)">
   </td>
  </tr>
 </form><% End If %>
 <tr> 
  <td valign="top"> 
   <table width="99%" border="0" cellspacing="1" cellpadding="0" bgcolor="B4E2EF" class="fonts">
	<form name="form" method="post" >
	 <tr bgcolor="F4F4F4"> 
	  <Td width="6%" align="center">&nbsp;</td>
	  <td width="39%"><font color="990033">&nbsp;�ļ�/�ļ����� </font></td>
	  <td width="13%" align="center"><font color="990033">�ļ����</font></td>
	  <td width="13%" align="center"><font color="990033">����</font></td>
	  <td width="15%" align="center"><font color="990033">�ļ���С</font></td>
	  <td width="27%" align="center"><font color="990033">����޸�ʱ��</font></td>
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
	  <td width="15%" align="center">��</td>
	  <td width="13%" align="center"><img src="<%=ImageFolder%>/ClosedFolder.gif" width="16" height="16" alt="�ļ���"></td>
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
		SWF�ļ�����Ԥ�� 
		<%elseif suffix=".asp" then%>
		<%else%>
		 <img src="<%=FileUrl(FileName,Dir)%>" alt="����鿴�ļ���ͼ" width="115" height="50">
		<%end if%>	   
	   </a></td>
	  <td width="13%" align="center"><Img src=<%=ImageFolder%>/<%=FileType%> alt="�ļ�"></td>
	  <td width="15%" align="center"><%=FileSize%></td>
	  <td width="27%" align="center"><%=FileTime%></td>
	 </tr>
	 <%	Next %>
	 <tr bgcolor="#FFFFFF"> 
	  <td width="6%" align="center"> 
	   <input type="checkbox" name="CheckAll" value="checkbox" onClick="CheckAll1()" title=ȫ��ѡ�� style="cursor:hand">	  </td>
	  <td colspan="5" height="30">
<%strHtml=Trim(Request.QueryString("action"))
 if strHtml="admin" then%>
	   <input type="button" name="Submit" value="�� ��" class=button style="cursor:hand" onClick="Edit()"  title=�༭>
	   <input type="button" name="Submit3" value="������" class=button style="cursor:hand" onClick="Rname()"  title=������>
	   <% End If %>
	   <input type="button" name="Submit2" value="ɾ ��" class=button style="cursor:hand" onClick="DelAll()"  title=ɾ��>
	   <font color="990033"> <input type="hidden" name="ThisDir" value="<%=Dir%>"></font>
	   <input name="button" type="button" class=button style="cursor:hand" onClick="window.location.href='admin_file.asp'"  value="��������ļ�"></td>
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