<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>无组件缩略图上传</title>
<style>
*{mrgin:0px;padding:0px;font-size:12px;}
form {margin:0px;padding:0px;}
form .inp{ width:140px;border:#999999 1px solid}
body{margin:0px;padding:0px; background:none}
</style>
</head>
<body>
<form action="upload.asp?formname=<%=request("formname")%>&inputname=<%=request("inputname")%>&uploadstyle=<%=request("uploadstyle")%>" method="post" enctype="multipart/form-data">
<input type="file" name="file1" class="inp" /> <input type="submit" class="btn2" value=" 上传 " />
</form>
</body>
</head>
</html>