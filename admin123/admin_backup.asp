<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->  
<%call chkAdmin(31)
action=trim(request("action"))
Dim zc_database_path
zc_database_path=""&db&""'数据库的路径
data_array= Split(zc_database_path,"/")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>数据库管理系统</title>
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">数据库管理系统</div></td>
</tr>
<tr>
<td bgcolor="#FFFFFF">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
<tr bgcolor="#FFFFFF" >
  <td height="25" align="center" class="td">
 <div class="Action">
<% 
If Action ="CompressData" Then c1="Action_1" else c1="Action_2"
If Action ="BackupData" Then c2="Action_1" else c2="Action_2"
If Action ="RestoreData" Then c3="Action_1" else c3="Action_2"
If Action ="SpaceSize" Then c4="Action_1" else c4="Action_2"
%>
<a href="?action=CompressData" class="<%=c1%>">[压缩数据库]</a> 
<a href="?action=BackupData" class="<%=c2%>">[备份数据库]</a> 
<a href="?action=RestoreData" class="<%=c3%>">[还原数据库]</a> 
<a href="?action=SpaceSize" class="<%=c4%>">[系统空间占用]</a>
 </div>
  </td>
</tr>
</table>

<%
Select Case action
Case "CompressData" '压缩数据
    Dim tmprs
    dim allarticle
    dim Maxid
    dim topic,username,dateandtime,body
    call CompressData()
	
case "BackupData" '备份数据
    if request("act")="Backup" Then
      call updata()
    else
      call BackupData()
    end If
case "RestoreData" '恢复数据
    dim backpath
    if request("act")="Restore" Then
      Dbpath=request.form("Dbpath")
      backpath=request.form("backpath")
       if dbpath="" Then
         response.write "请输入您的数据库的全名!" 
       else
         Dbpath=server.mappath(Dbpath)
       end If
      backpath=server.mappath(backpath)
      Set Fso=server.CreateObject("scripting.filesystemobject")
      if fso.fileexists(dbpath) Then 
        fso.copyfile Dbpath,Backpath
        response.write "<div class=""d"">数据库被成功还原!</div>"
      else
        response.write "没找到您所需要的数据库!" 
      end If
    else
      call RestoreData()
    end If
Case "SpaceSize" '系统空间占用
    call SpaceSize()
Case "deletebackup"
    Dim dbname
    dbpath=Request.QueryString("dbpath")
    dbname=Request.QueryString("dbname")
    dbpath=Server.MapPath(dbpath)
    dbpath=dbpath &"\"&dbname
    set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(dbPath) Then
      fso.DeleteFile(DBPath)
      Set fso = nothing
      response.write "<div class=""d"">您备份的数据库已经" & dbpath &"被成功删除!</div>"
    Else
      response.write dbpath 
      response.write "<div class=""d"">输入的路径错误,请确认后重新输入!</div>"
    End If
Case Else
End Select%>
</td>
</tr>
</table>

<%Sub SpaceSize()'====================系统空间占用=======================
On Error Resume Next%>
<div align="center" style="line-height:30px">数据库空间使用情况 数据库:<font color="#FF0000"><%showSpaceinfo("../"&data_array(1)&"")%></font>  备份数据库:<font color="#FF0000"><%showSpaceinfo("databackup")%></font> 系统总共:<font color="#FF0000"><%showSpaceinfo("/")%></font></div>
<%End Sub%>

<%Sub ShowSpaceInfo(drvpath)
dim fso,d,size,showsize
set fso=server.CreateObject("scripting.filesystemobject") 
drvpath=server.mappath(drvpath) 
set d=fso.getfolder(drvpath) 
size=d.size
showsize=size & " Byte" 
if size>1024 Then
size=(Size/1024)
showsize=size & " KB"
end If
if size>1024 Then
size=(size/1024)
showsize=formatnumber(size,2) & " MB" 
end If
if size>1024 Then
size=(size/1024)
showsize=formatnumber(size,2) & " GB" 
end If 
response.write "<font face=verdana>" & showsize & "</font>"
End Sub %>

<%Sub RestoreData()%>
<table width="100%" border="0"  cellpadding="5" cellspacing="0" class="stable">
<form id="edit" method="post" action="?action=RestoreData&act=Restore">
<tr>
    <td width="310" align="right" class="td">还原的路径(相对路径):</td>
    <td class="td"><input type=text size=30 name=DBpath value="Databackup/Databack.mdb">(默认的，可手动修改保持路径正确)</td>
  </tr>
<tr>
    <td align="right" class="td">还原的路径(相对路径):</td>
    <td class="td"><input type=text size=30 name=backpath value="<%=zc_database_path%>"  readonly style="background:#CCC"> <input type=submit class="btn2" value="开始还原"></td>
  </tr>
</form>
</table>
<%End Sub%>

<%Sub updata()
Dbpath=request.form("Dbpath")
Dbpath=server.mappath(Dbpath)
bkfolder=request.form("bkfolder")
bkdbname=request.form("bkdbname")
Set Fso=server.CreateObject("scripting.filesystemobject")
if fso.fileexists(dbpath) Then
    If CheckDir(bkfolder) = True Then
      'response.Write Dbpath
      'response.End 
      fso.copyfile Dbpath,bkfolder& "\"& bkdbname
    else
      MakeNewsDir bkfolder
      fso.copyfile dbpath,bkfolder& "\"& bkdbname
    end If
    response.write "<div class=""d"">已经成功备份,你的数据库的路径:" &bkfolder& "\"& bkdbname&"</div>"
    response.write "<div class=""d"">点击此处将数据库下载下来:<a href="""& ZC_BLOG_HOST &request.form("bkfolder") & "/" & bkdbname &""">" & ZC_BLOG_HOST & request.form("bkfolder") & "/" & bkdbname &"</div>"
    response.write "<div class=""d""><a href=""?action=deletebackup&dbpath="&request.form("bkfolder") &"&dbname=" & bkdbname &""">当您下载完毕后,点击此处将删除备份的数据库!</a></div>"
   Else
    response.write "<div class=""d"">Error ,,找不到文件!</div>"
End If
Set fso = nothing
End Sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
folderpath=Server.MapPath(".")&"\"&folderpath
Set fso1 = CreateObject("Scripting.FileSystemObject")
If fso1.FolderExists(FolderPath) Then
    '存在
    CheckDir = True
Else
    '不存在
    CheckDir = False
End If
Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
dim f
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set f = fso1.CreateFolder(foldername)
MakeNewsDir = True
Set fso1 = nothing
End Function%>
  
<%Sub BackupData()%>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<form id="edit" method="post" action="?action=BackupData&act=Backup">
  <tr>
    <td width="310" align="right" class="td">当前数据库的路径(相对路径):</td>
    <td class="td"><input type=text size=30 name=DBpath value="<%=zc_database_path%>" readonly style="background:#CCC"></td>
  </tr>
  <tr>
    <td align="right" class="td">备份数据库的路径(相对路径):</td>
    <td class="td"><input name=bkfolder type=text value="Databackup" size=30 readonly="readonly">如果该目录不存在,系统将自动建立</td>
  </tr>
  <tr>
    <td align="right" class="td">备份后数据库的名称:</td>
    <td class="td"><input type=text size=30 name=bkDBname readonly style="background:#CCC" value="Databack.mdb">       
如果备份文件不存在将建立,如果存在,将自动覆盖!</td>
  </tr>
  <tr>
    <td align="right" class="td">&nbsp;</td>
    <td class="td"><input type=submit class="btn2" value="开始备份"></td>
  </tr>
</form>
</table>
<%End Sub%>

<%Sub CompressData()%>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<form id="edit" action="?action=CompressData" method="post">
  <tr>
    <td width="310" align="right" class="td">输入数据库的所在路径:</td>
    <td class="td"><input name="dbpath" type="text" value="<%=zc_database_path%>"  readonly style="background:#CCC" size="30">
      <input type="checkbox" name="boolIs97" value="True">
      如果是access97,请将钩打上.(默认是Access 2000)</td>
  </tr>
    <tr>
    <td align="right" class="td">&nbsp;</td>
    <td class="td"><input type="submit" class="btn2" value="开始压缩"></td>
  </tr>
</form>
</table>
<%dbpath = request("dbpath")
boolIs97 = request("boolIs97")
If dbpath <> "" Then
dbpath = server.mappath(dbpath)
response.write(CompactDB(dbpath,boolIs97))
End If
End Sub%>

<%
'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = Left(dbPath,InStrRev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
    fso.CopyFile dbpath,strDBPath & "temp.mdb"
    Set Engine = CreateObject("JRO.JetEngine")
     If boolIs97 = "True" Then
      Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
& "Jet OLEDB:Engine Type=" & JET_3X
    Else
      Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
    End If
    fso.CopyFile strDBPath & "temp1.mdb",dbpath
    fso.DeleteFile(strDBPath & "temp.mdb")
    fso.DeleteFile(strDBPath & "temp1.mdb")
    Set fso = nothing
    Set Engine = nothing
    CompactDB = "<div class=""d"">您的数据库" & dbpath & "已经被成功压缩!</div>" & vbCrLf
Else
    CompactDB = "<div class=""d"">您输入的路径错误,请确认后重新输入!</div>" & vbCrLf
End If
End Function%>

</body>
</html>