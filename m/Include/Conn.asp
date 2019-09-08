<%
response.buffer=true '启用缓冲处理

dir="/"'------------根目录为"/",二级目录时请填目录名，如："/domo/"
db=""&dir&"data/zychcms.vip#flagship@#@#!.mdb"'数据库链接路径
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.open connstr
Call OpenConn()
If Err Then
err.Clear
Set Conn = Nothing
Response.Write "数据库连接出错，请检查数据库连接文件中的数据库参数设置。"
Response.End
End If 
Private Sub OpenConn()
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnStr
End Sub
sub Chkhttp()
server_vv=len(Request.ServerVariables("SERVER_NAME"))
server_v1=left(Cstr(Request.ServerVariables("HTTP_REFERER")),server_vv)
server_v2=left(Cstr("http://"&Request.ServerVariables("SERVER_NAME")),server_vv)
if server_v1<>server_v2 or server_v1="" or server_v1="" then
response.Charset="utf-8"
response.write("<script>alert('错误：禁止从站点外部提交数据!.')</script>")
response.end
end if
end sub
Private Sub zych_Nav()
End Sub

Function Getname(str,shu)
filename=str 
filename=split(filename,"/")
Getname=filename(shu)
End Function
htmlDir1="" 
upLoadPath="Uploadfile/image"
upFileSize=20971520 '最大上传大小，默认是2M
upFileWay=1 	'1:按天存入目录 2:按月存入目录 3:按扩展名存目录
upFiletype="txt,rar,zip,doc,xls,ppt,jpg,jpeg,gif,png,swf,wmv,avi,wma,mp3,mid,flv"'上传扩展名
'函数名：ThumbnailImg
'作 用：制作图片的缩略图
'参 数：ImgUrl 原图地址
' ImgWidth 新图的宽
' Imgheight 新图的高
' newImgurl 新图的存放地址
'****************************************************
Sub ThumbnailImg(Imgurl,ImgWidth,Imgheight,newImgurl)
Dim Jpeg ''''//声明变量
If instr(imgUrl,":\")=0 Then imgUrl = Server.MapPath(imgUrl) 
If instr(newImgurl,":\")=0 Then newImgurl = Server.MapPath(newImgurl) 
Set Jpeg = Server.CreateObject("Persits.jpeg") ''''//调用组件
Jpeg.Open Imgurl ''''//原图位置
If ImgWidth<>"" and Imgheight="" Then Imgheight = ImgWidth*Jpeg.OriginalHeight / Jpeg.OriginalWidth 
If Imgheight<>"" and ImgWidth="" Then ImgWidth = Imgheight*Jpeg.OriginalWidth / Jpeg.OriginalHeight 
Jpeg.Width = ImgWidth ''''//设图片宽度
Jpeg.Height = Imgheight ''''//设图片高度
Jpeg.Sharpen 1, 130 ''''//设定锐化效果
Jpeg.Save newImgurl ''''//生成缩略图位置及名称
Set Jpeg = Nothing ''''//注销组件，释放资源
End Sub
%>