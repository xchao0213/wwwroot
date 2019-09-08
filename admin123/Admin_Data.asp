<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(8)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站后台管理</title>
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">程序占用空间情况</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#6298E1" class="stable">
  <tr>
    <td width="19%" height="30" align="right" class="td">系统占用空间总计：</td>
    <td width="81%" class="td"><img src="images/bg_list.gif" width=<%=drawbar("allsize")%> height=9 />&nbsp;
    <%allsize()%></td>
  </tr>
  <tr>
    <td height="30" align="right" class="td">数据库占用空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("data")%> height=9 />&nbsp;
    <%othersize("data")%></td>
  <tr>
    <td height="30" align="right" class="td">系统后台占用空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("Admin")%> height=9 />&nbsp;
    <%othersize("Admin")%></td>
  </tr>
  <tr>
    <td height="30" align="right" class="td">备份数据占用空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("admin/Databackup")%> height=9 />&nbsp;
    <%othersize("admin/Databackup")%></td>
  </tr>
  <tr>
    <td height="30" align="right" class="td">系统图片目录占用空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("images")%> height=9 />&nbsp;
    <%othersize("images")%></td>
  </tr>
    <tr>
    <td height="30" align="right" class="td">编辑器占用空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("zycheditor")%> height=9 />&nbsp;
    <%othersize("zycheditor")%></td>
  </tr>
    <tr>
    <td height="30" align="right" class="td">图片空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("images")%> height=9 />&nbsp;
    <%othersize("images")%></td>
  </tr>
  <tr>
    <td height="30" align="right" class="td">上传目录占用空间：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("uploadfile")%> height=9 />&nbsp;
    <%othersize("uploadfile")%> <a href="admin_Attachment.asp" target="main"><strong style="color:#FF0000">附件管理</strong></a></td>
  </tr>
    <tr>
    <td height="30" align="right" class="td">支付宝返回日志：</td>
    <td class="td"><img src="images/bg_list.gif" width=<%=drawbar("alipay")%> height=9 />&nbsp;
    <%othersize("alipay/Guarantee/log")%> <a href="admin_alipay_log.asp" target="main"><strong style="color:#FF0000">日志管理</strong></a></td>
  </tr>
</table></td>
  </tr></form>
</table>

<%
sub othersize(names)
	dim fso,path,ml,mlsize,dx,d,size
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	path=server.mappath("..\Images")
	ml=left(path,(instrrev(path,"\")-1))&"\"&names
	
	On Error Resume Next
	set d=fso.getfolder(ml) 
	If Err Then
		err.Clear
		Response.Write "<font color=red>提示：没有“"&names&"”目录</font>"					
		'Response.End()
	End If
	mlsize=d.size
	size=mlsize
	dx=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   dx=formatnumber(size,2) & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write dx
end sub

sub allsize()
	dim fso,path,ml,mlsize,dx,d,size
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	path=server.mappath("../index.asp")
	ml=left(path,(instrrev(path,"\")-1))
	set d=fso.getfolder(ml) 
	mlsize=d.size
	size=mlsize
	dx=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   dx=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write dx
end sub

Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	drvpathroot=server.mappath("../Images")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	On Error Resume Next
	drvpath=server.mappath("../"&drvpath)
	If Err Then
		err.Clear
		Response.Write "没有名为“"&drvpath&"”的目录，您可以修改文件以正确显示该目录的使用量。"			
		Response.End()
	End If
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 

Function FileList(FolderUrl,FileExName)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set folder=fso.GetFolder(Server.MapPath(Trim(FolderUrl)))
Set file=folder.Files
FileList=""
For Each FileName in file
If Trim(FileExName)<>"" Then
	If InStr(Trim(FileExName),Trim(Mid(FileName.Name,InStr(FileName.Name,".")+1,len(FileName.Name))))>0 Then
    	FileList=FileList&""&FileName.Name&"|"
	End If
Else
     FileList=FileList&"<a href='#'>"&FileName.Name&"</a><br>"
End If
Next
Set file=Nothing
Set folder=Nothing
Set fso=Nothing
End Function
%>
</body>
</html>