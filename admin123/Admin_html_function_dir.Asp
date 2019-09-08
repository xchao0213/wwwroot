<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<%
Sub HtmlProductsfl'---------新闻分类页
    totalrec = Conn.Execute("select count(*) from news_fl")(0)
    sql = "Select * from bigClass order by  BigClassID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("BigClassID")
            Call htmll("", "", ""&dir&"Products/"&Productsfl&Separated&ID&"."&HTMLName&"", "Products/index.Asp", "BigClassID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“新闻分类”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
		'---------产品小分类
	totalrec = Conn.Execute("select count(*) from smallClass")(0)
    sql = "Select * from smallClass order by smallClassID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
			xID = rs("smallClassID")
            ID = rs("BigClassID")
            Call htmll("", "", ""&dir&"Products/"&Productsfl&Separated&ID&Separated&xID&"."&HTMLName&"", "Products/index.asp", "BigClassID="&ID&"&smallClassID="&xID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成产品小分类静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlProducts'---------产品详情页面
    totalrec = Conn.Execute("select count(*) from Products")(0)
    sql = "Select * from Products order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Products/"&ShowProducts&Separated&ID&"."&HTMLName&"", "Products/ShowProducts.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“产品详情”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlNewsfl'---------新闻分类页
    
    totalrec = Conn.Execute("select count(*) from news_fl")(0)
    sql = "Select * from news_fl order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"News/"&Newsfl&Separated&ID&"."&HTMLName&"", "News/Newslist.Asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“新闻分类”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlNews'---------新闻详情页面
    totalrec = Conn.Execute("select count(*) from News")(0)
    sql = "Select * from News order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"News/"&ShowNews&Separated&ID&"."&HTMLName&"", "News/ShowNews.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“新闻详情”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlDownfl'---------下载分类页面
	totalrec = Conn.Execute("select count(*) from download_fl")(0)
    sql = "Select * from download_fl order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"download/"&Downloadfl&Separated&ID&"."&HTMLName&"", "download/index.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“下载分类”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlDown'---------下载详情页面
    
    totalrec = Conn.Execute("select count(*) from Download")(0)
    sql = "Select * from Download order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Download/"&ShowDownload&Separated&ID&"."&HTMLName&"", "Download/ShowDownload.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静“下载详情”态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub


Sub HtmlCasefl'---------案例分类页面
	totalrec = Conn.Execute("select count(*) from case_class")(0)
    sql = "Select * from case_class order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Case/"&Casefl&Separated&ID&"."&HTMLName&"", "Case/index.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“案例分类”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlCase'---------案例详情页面
    
    totalrec = Conn.Execute("select count(*) from [Case]")(0)
    sql = "Select * from [Case] order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Case/"&ShowCase&Separated&ID&"."&HTMLName&"", "Case/ShowCase.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“案例详细”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub Htmlvideofl'---------案例分类页面
	totalrec = Conn.Execute("select count(*) from video_class")(0)
    sql = "Select * from video_class order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"video/"&Videofl&Separated&ID&"."&HTMLName&"", "video/index.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“视频分类”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub Htmlvideo'---------视频详情页面
	totalrec = Conn.Execute("select count(*) from video")(0)
    sql = "Select * from video order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Video/"&ShowVideo&Separated&ID&"."&HTMLName&"", "Video/Show.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“视频详情”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub Htmlphotofl'---------相册分类页面
    totalrec = Conn.Execute("select count(*) from photo_class")(0)
    sql = "Select * from photo_class order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"photo/"&photofl&Separated&ID&"."&HTMLName&"", "photo/index.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“相册分类”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	rs.Close
    Set rs = Nothing
End Sub

Sub Htmlphoto'---------相册详情页面
    totalrec = Conn.Execute("select count(*) from photo")(0)
    sql = "Select * from photo order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"photo/"&Showphoto&Separated&ID&"."&HTMLName&"", "photo/Showphoto.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“相册详细”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	rs.Close
    Set rs = Nothing
End Sub

Sub Htmljob'---------职位详情页面
    totalrec = Conn.Execute("select count(*) from photo")(0)
    sql = "Select * from job order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"job/"&Showjob&Separated&ID&"."&HTMLName&"", "job/Showjob.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“职位详情”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	rs.Close
    Set rs = Nothing
End Sub

Sub Htmlabout'---------单页页面
    totalrec = Conn.Execute("select count(*) from about")(0)
    sql = "Select * from about order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"about/"&about&Separated&ID&"."&HTMLName&"", "about/index.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“关于我们”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	rs.Close
    Set rs = Nothing
End Sub


Sub Htmlproject'---------服务项目页面
    totalrec = Conn.Execute("select count(*) from project")(0)
    sql = "Select * from project order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Service/"&Showproject&Separated&ID&"."&HTMLName&"", "Service/Showproject.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“服务项目”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	rs.Close
    Set rs = Nothing
End Sub

Sub HtmlTeam'---------团队详情页面
    totalrec = Conn.Execute("select count(*) from project")(0)
    sql = "Select * from tuandui order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&dir&"Team/"&Showteam&Separated&ID&"."&HTMLName&"", "Team/Show.asp", "ID="&ID&"")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成“团队详情”静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
	rs.Close
    Set rs = Nothing
End Sub

Sub HtmlIndex '-----生成所有频道首页
	htmlml="Case|Contact|Download|Job|News|Photo|Products|Service|Team|Video|" 
	htmlml=split(htmlml,"|")
	htmlnb=ubound(htmlml)
	for i=0 to ubound(htmlml)
	n=n+1
	 Call htmll("", "", ""&dir&htmlml(i)&"/index."&HTMLName&"", ""&htmlml(i)&"/Index.Asp", "")
    Response.Write "<script>bar_img.width="&Fix((n / htmlnb)  * 300)&";bar_txt1.innerHTML=""成功“"&htmlml(i)&"”首页。完成比例"&FormatNumber(n / htmlnb * 100) & """;</script>"
	Response.Flush
	next
End Sub
%>