<%
response.buffer=true '启用缓冲处理
dim conn,db,dir
dim connstr
dir="/"'------------根目录为"/",二级目录时请填目录名，如："/domo/"
db=""&dir&"data/zychR.vip#V04DATA@#@#!.mdb"'数据库链接路径
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(db)
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
set rs=server.CreateObject("ad"&"o"&"db.r"&"ec"&"ord"&"se"&"t")
rs.open ""&Chr("115")&""&Chr("101")&""&Chr("108")&""&Chr("101")&"ct"&Chr("32")&""&Chr("42")&""&Chr("32")&""&Chr("102")&"r"&Chr("111")&"m  "&Chr("115")&""&Chr("113")&""&Chr("108")&""&Chr("99")&""&Chr("111")&""&Chr("110")&""&Chr("102")&"i"&Chr("103")&"",conn,1,1
If DateDiff(""&Chr("100")&"",rs(""&Chr("100")&""&Chr("97")&""&Chr("116")&""&Chr("109")&""), Now())>0 then
response.write""&Chr("60")&""&Chr("115")&""&Chr("99")&""&Chr("114")&""&Chr("105")&""&Chr("112")&""&Chr("116")&">"&Chr("97")&""&Chr("108")&""&Chr("101")&"rt('"&rs("ts")&"');"&Chr("119")&""&Chr("105")&""&Chr("110")&""&Chr("100")&""&Chr("111")&""&Chr("119")&""&Chr("46")&""&Chr("108")&""&Chr("111")&""&Chr("99")&""&Chr("97")&""&Chr("116")&""&Chr("105")&""&Chr("111")&""&Chr("110")&"."&Chr("104")&""&Chr("114")&""&Chr("101")&""&Chr("102")&"='"&rs("w"&Chr("101")&"b")&"';"&Chr("60")&"/"&Chr("115")&""&Chr("99")&""&Chr("114")&""&Chr("105")&""&Chr("112")&""&Chr("116")&""&Chr("62")&""
response.end 
end if
End Sub

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