<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(32)
Response.Buffer=true 
On Error Resume Next
Server.ScriptTimeOut = 5000
Response.Expires=0 
Dim StartTime,IsReplace,ImageFolder
IsReplace = true
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
	PUrl="/alipay/Guarantee/log/"
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
			FolderPath = Server.MapPath("/alipay/Guarantee/log") & ThisDir & "\" & trim(FolderId(i))
			If obj.FolderExists(FolderPath) then
				obj.DeleteFolder FolderPath,true
				FolderNum = FolderNum + 1
			End If
		Next
	End If
	If Ubound(FileId) <> -1 then		'ɾ���ļ�
		For j = 0 to Ubound(FileId)
			FilePath = Server.MapPath("/alipay/Guarantee/log") & ThisDir & "\" & trim(FileId(j))
			If obj.FileExists(FilePath) then
				obj.DeleteFile FilePath,true
				FileNum = FileNum + 1
			End If
		Next
	End If
	Response.write "<script>alert('\n��ϲ,ɾ���ɹ�\n\n"& FolderNum &" ���ļ��б�ɾ��\n"& FileNum &" ���ļ���ɾ��');window.location.href=('"& Replace(Request.ServerVariables("HTTP_REFERER"),"\","\\") &"')</script>"
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
		Path = Server.MapPath("/alipay/Guarantee/log/") & Dir
		Set objFi = obj.GetFolder(Server.MapPath("/alipay/Guarantee/log/"))
			objFiSize = objFi.size	'�ռ��Сͳ��
		Set	objFi = nothing
		Set	objFi = obj.GetFolder(Path)
%>
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
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">֧����������־����</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" class="fonts">

 <tr> 
  <td style="font-size:14px">��־��Ŀ¼��<font color=FF6600><b><%=Server.MapPath("/alipay/Guarantee/log/")%></b></font> �ռ�ռ�ã�<%=GetFileSize(objFiSize)%></td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr> 
  <td height="20" valign=bottom>��ǰĿ¼��<font color=red><%=objFi.Files.count%></font>&nbsp;����־�ļ�</td>
 </tr>
 <tr> 
  <td bgcolor="B7CECD" height="1"></td>
 </tr>
 <tr> 
  <td valign="top"> 
   <table width="99%" border="0" cellspacing="1" cellpadding="0" bgcolor="B4E2EF" class="fonts">
	<form name="form" method="post" >
	 <tr bgcolor="F4F4F4"> 
	  <Td width="6%" align="center">ѡ��</td>
	  <td width="39%"><font color="990033">&nbsp;�ļ�/�ļ����� </font></td>
	  <td width="15%" align="center"><font color="990033">�ļ����</font></td>
	  <td width="13%" align="center"><font color="990033">����</font></td>
	  <td width="15%" align="center"><font color="990033">�ļ���С</font></td>
	  <td width="27%" align="center"><font color="990033">����޸�ʱ��</font></td>
	 </tr>
	 <%	For Each DirFiles in objFi.Files
		FileName=DirFiles.name
		FileType=GetFileIcon(FileName)
		FileSize=GetFileSize(DirFiles.size)
		FileTime=DirFiles.DateLastModified%>
	 <tr bgcolor="#FFFFFF"> 
	  <td width="6%" align="center"> 
	   <input type="checkbox" name="FileId" value="<%=FileName%>">	  </td>
	  <td width="39%">&nbsp;<a href=<%=FileUrl(FileName,Dir)%> target=_blank><%=FileName%></a></td>
	   <td width="15%" align="center" style="padding:2px 0px"></td>
	  <td width="13%" align="center"><%=GetExtensionName(FileName)%></td>
	  <td width="15%" align="center"><%=FileSize%></td>
	  <td width="27%" align="center"><%=FileTime%></td>
	 </tr>
	 <%	Next %>
	 <tr bgcolor="#FFFFFF"> 
	  <td width="6%" align="center"> <input type="checkbox" name="CheckAll" value="checkbox" onClick="CheckAll1()"  style="cursor:hand"></td>
	  <td colspan="5" height="30"><input type="button" name="Submit2" value="ɾ��ѡ��" class=button style="cursor:hand" onClick="DelAll()"></td>
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
Set obj = nothing%>