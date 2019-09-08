<!--#include file="Admin_html_function_dir.Asp"-->
<%
server.scripttimeout = 999

Function htmll(mulu, htmlmulu, FileName, filefrom, htmla)
    If mulu = "" Then mulu = ""&Dir&""
    If htmlmulu = "" Then htmlmulu = ""&Dir&""
    mulu = Replace(mulu,"//", "/")
    FilePath = Server.MapPath(mulu)&"\"&FileName
    Do_Url = "http://"
    Dirs = "" '------------根目录为空，二级目录时请填目录名，如："/domo/"
    Do_Url = Do_Url&Request.ServerVariables("server_name")&htmlmulu&""&dirs&""&filefrom
    Do_Url = Do_Url&"?"&htmla
    strUrl = Do_Url
    Set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
    objXmlHttp.Open "GET", strUrl, False
    objXmlHttp.send()
    binFileData = objXmlHttp.responseBody
    Set objXmlHttp = Nothing
    Set objAdoStream = Server.CreateObject("Adodb." & "Stream")
    objAdoStream.Type = 1
    objAdoStream.Open()
    objAdoStream.Write(binFileData)
    objAdoStream.SaveToFile FilePath, 2
    objAdoStream.Close()
    Set objAdoStream = Nothing
End Function
dim objfso,path
set objfso=Server.CreateObject("Scripting.FileSystemObject") 
SourceFolder =server.MapPath(""&Dir&""&HTMLdir&"") 
'判断文件夹是否存在,如果不存在则创建文件夹
If not objfso.FolderExists(SourceFolder) then
 objfso.CreateFolder SourceFolder    
End if
set objfso=nothing
%>
