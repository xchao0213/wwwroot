<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->  
<%call chkAdmin(31)
action=trim(request("action"))
Dim zc_database_path
zc_database_path=""&db&""'���ݿ��·��
data_array= Split(zc_database_path,"/")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ݿ����ϵͳ</title>
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<tr>
<td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">���ݿ����ϵͳ</div></td>
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
<a href="?action=CompressData" class="<%=c1%>">[ѹ�����ݿ�]</a> 
<a href="?action=BackupData" class="<%=c2%>">[�������ݿ�]</a> 
<a href="?action=RestoreData" class="<%=c3%>">[��ԭ���ݿ�]</a> 
<a href="?action=SpaceSize" class="<%=c4%>">[ϵͳ�ռ�ռ��]</a>
 </div>
  </td>
</tr>
</table>

<%
Select Case action
Case "CompressData" 'ѹ������
    Dim tmprs
    dim allarticle
    dim Maxid
    dim topic,username,dateandtime,body
    call CompressData()
	
case "BackupData" '��������
    if request("act")="Backup" Then
      call updata()
    else
      call BackupData()
    end If
case "RestoreData" '�ָ�����
    dim backpath
    if request("act")="Restore" Then
      Dbpath=request.form("Dbpath")
      backpath=request.form("backpath")
       if dbpath="" Then
         response.write "�������������ݿ��ȫ��!" 
       else
         Dbpath=server.mappath(Dbpath)
       end If
      backpath=server.mappath(backpath)
      Set Fso=server.CreateObject("scripting.filesystemobject")
      if fso.fileexists(dbpath) Then 
        fso.copyfile Dbpath,Backpath
        response.write "<div class=""d"">���ݿⱻ�ɹ���ԭ!</div>"
      else
        response.write "û�ҵ�������Ҫ�����ݿ�!" 
      end If
    else
      call RestoreData()
    end If
Case "SpaceSize" 'ϵͳ�ռ�ռ��
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
      response.write "<div class=""d"">�����ݵ����ݿ��Ѿ�" & dbpath &"���ɹ�ɾ��!</div>"
    Else
      response.write dbpath 
      response.write "<div class=""d"">�����·������,��ȷ�Ϻ���������!</div>"
    End If
Case Else
End Select%>
</td>
</tr>
</table>

<%Sub SpaceSize()'====================ϵͳ�ռ�ռ��=======================
On Error Resume Next%>
<div align="center" style="line-height:30px">���ݿ�ռ�ʹ����� ���ݿ�:<font color="#FF0000"><%showSpaceinfo("../"&data_array(1)&"")%></font>  �������ݿ�:<font color="#FF0000"><%showSpaceinfo("databackup")%></font> ϵͳ�ܹ�:<font color="#FF0000"><%showSpaceinfo("/")%></font></div>
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
    <td width="310" align="right" class="td">��ԭ��·��(���·��):</td>
    <td class="td"><input type=text size=30 name=DBpath value="Databackup/Databack.mdb">(Ĭ�ϵģ����ֶ��޸ı���·����ȷ)</td>
  </tr>
<tr>
    <td align="right" class="td">��ԭ��·��(���·��):</td>
    <td class="td"><input type=text size=30 name=backpath value="<%=zc_database_path%>"  readonly style="background:#CCC"> <input type=submit class="btn2" value="��ʼ��ԭ"></td>
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
    response.write "<div class=""d"">�Ѿ��ɹ�����,������ݿ��·��:" &bkfolder& "\"& bkdbname&"</div>"
    response.write "<div class=""d"">����˴������ݿ���������:<a href="""& ZC_BLOG_HOST &request.form("bkfolder") & "/" & bkdbname &""">" & ZC_BLOG_HOST & request.form("bkfolder") & "/" & bkdbname &"</div>"
    response.write "<div class=""d""><a href=""?action=deletebackup&dbpath="&request.form("bkfolder") &"&dbname=" & bkdbname &""">����������Ϻ�,����˴���ɾ�����ݵ����ݿ�!</a></div>"
   Else
    response.write "<div class=""d"">Error ,,�Ҳ����ļ�!</div>"
End If
Set fso = nothing
End Sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
folderpath=Server.MapPath(".")&"\"&folderpath
Set fso1 = CreateObject("Scripting.FileSystemObject")
If fso1.FolderExists(FolderPath) Then
    '����
    CheckDir = True
Else
    '������
    CheckDir = False
End If
Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
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
    <td width="310" align="right" class="td">��ǰ���ݿ��·��(���·��):</td>
    <td class="td"><input type=text size=30 name=DBpath value="<%=zc_database_path%>" readonly style="background:#CCC"></td>
  </tr>
  <tr>
    <td align="right" class="td">�������ݿ��·��(���·��):</td>
    <td class="td"><input name=bkfolder type=text value="Databackup" size=30 readonly="readonly">�����Ŀ¼������,ϵͳ���Զ�����</td>
  </tr>
  <tr>
    <td align="right" class="td">���ݺ����ݿ������:</td>
    <td class="td"><input type=text size=30 name=bkDBname readonly style="background:#CCC" value="Databack.mdb">       
��������ļ������ڽ�����,�������,���Զ�����!</td>
  </tr>
  <tr>
    <td align="right" class="td">&nbsp;</td>
    <td class="td"><input type=submit class="btn2" value="��ʼ����"></td>
  </tr>
</form>
</table>
<%End Sub%>

<%Sub CompressData()%>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<form id="edit" action="?action=CompressData" method="post">
  <tr>
    <td width="310" align="right" class="td">�������ݿ������·��:</td>
    <td class="td"><input name="dbpath" type="text" value="<%=zc_database_path%>"  readonly style="background:#CCC" size="30">
      <input type="checkbox" name="boolIs97" value="True">
      �����access97,�뽫������.(Ĭ����Access 2000)</td>
  </tr>
    <tr>
    <td align="right" class="td">&nbsp;</td>
    <td class="td"><input type="submit" class="btn2" value="��ʼѹ��"></td>
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
'=====================ѹ������=========================
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
    CompactDB = "<div class=""d"">�������ݿ�" & dbpath & "�Ѿ����ɹ�ѹ��!</div>" & vbCrLf
Else
    CompactDB = "<div class=""d"">�������·������,��ȷ�Ϻ���������!</div>" & vbCrLf
End If
End Function%>

</body>
</html>