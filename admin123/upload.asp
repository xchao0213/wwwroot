<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/inc.asp" -->
<!--#include file="seeion.asp"-->
<!--#include file="upLoad_Class.asp"-->
<html>
<head>
<title>无组件缩略图上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" type=text/css rel=stylesheet>
<style>
*{mrgin:0px;padding:0px;font-size:12px;}
body{margin:4px;color:#f00;}
body .btn2{font-size:12px;background: url(images/button_bg.gif);border:1px solid #BDC5CA; padding:2px 6px;  height:20px;color:#333333;cursor:pointer}
</style>
</head>
<body>
<%
formname=request("formname")
inputname=request("inputname")
uploadstyle=request("uploadstyle")
dim upload
set upload = new AnUpLoad
upload.Exe = "jpg|bmp|jpeg|gif|png|rar|zip|ppt|doc|swf|apk|ios"
upload.MaxSize = 10 * 1024 * 1024 '2M
upload.GetData()
if upload.ErrorID>0 then 
	response.Write upload.Description
else
	dim file,savpath
	If uploadstyle="Logo" Then 
	savepath = "/uploadfile/image/Logo"
	ElseIf uploadstyle="News" Then 
	savepath = "/uploadfile/image/News"
	ElseIf uploadstyle="Culture" Then 
	savepath = "/uploadfile/image/Culture"
	ElseIf uploadstyle="case" Then 
	savepath = "/uploadfile/image/case"
	ElseIf uploadstyle="Products" Then 
	savepath = "/uploadfile/image/Products"
	ElseIf uploadstyle="Download" Then 
	savepath = "/uploadfile/image/Download"
	ElseIf uploadstyle="Ad" Then 
	savepath = "/uploadfile/image/Ad"
	ElseIf uploadstyle="Service" Then 
	savepath = "/uploadfile/image/Service"
	ElseIf uploadstyle="tuandui" Then 
	savepath = "/uploadfile/image/tuandui"
	ElseIf uploadstyle="photo" Then 
	savepath = "/uploadfile/image/photo"
	ElseIf uploadstyle="video" Then 
	savepath = "/uploadfile/image/video"
	ElseIf uploadstyle="file" Then 
	savepath = "/uploadfile/file"
	End If 
	set file = upload.files("file1")
	if not(file is nothing) then
		set result = file.saveToFile(savepath,0,true)
		if not result.error Then
		    jrfilename=file.filename
            '水印代码位置 
			response.Write "上传成功,SIZE:" & file.size & " 字节"
			jrfilename=file.filename
            Response.write"<script language=javascript>parent.document."&formname&"."&inputname&".value='"&savepath&"/"&jrfilename&"';</script>"
			target=""&savepath&"/"&jrfilename&""
			%>
<script type="text/javascript">
function doInterfaceUpload(strValue){
	if (strValue=="") return;
	var objLinkUpload = parent.document.getElementsByName("ImagePath")[0];
	if (objLinkUpload){
		if (objLinkUpload.value!=""){
			objLinkUpload.value = objLinkUpload.value + "|";
		}
		objLinkUpload.value = objLinkUpload.value + strValue;
//		objLinkUpload.fireEvent("onchange");
	}
}
doInterfaceUpload('<%=target%>');

doChange(parent.document.getElementById("ImagePath"),parent.document.getElementById("ImageFileList"))

function doChange(objText, objDrop){
if (!objDrop) return;
var str = objText.value;
var arr = str.split("|");
var nIndex = objDrop.selectedIndex;
objDrop.length=1;
for (var i=0; i<arr.length; i++){
objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
}
objDrop.selectedIndex = nIndex;
}
</script>
<script type="text/javascript">
function divMsg(t){
  //alert(parent.document.getElementById("pw").innerHTML);
  var pstr=parent.document.getElementById("pw").innerHTML; 
   
  var imgvalue=parent.document.getElementById("ImagePath").value;

  if(t.indexOf("err:")==-1)
  {
  	var ids=t
	ids=ids.substr(ids.lastIndexOf("/")+1,ids.lastIndexOf("."))
	ids=ids.replace(".","")
  	 var itml = pstr + '<div class="imgDiv" id="'+ids+'"><a href="#" onclick="setimg(\''+t+'\')" value="'+t+'"><img src="'+ t + '" border="0" /></a><br><input type="radio" name="imgradio" onclick="setimg(\''+t+'\')" value="'+t+'" />设为缩略图 <a href="javascript:dropThisDiv(\''+ids+'\',\''+t+'\')">删除</a></div> ';
	 imgvalue=imgvalue+t+'|';
  }
  else
 {
 	 var itml = pstr + '<div class="imgDiv" id="'+t+'">"'+t+'"<br/>图片上传失败</div> ';
 } 
//document.getElementById("ImagePath").value=imgvalue;
 parent.document.getElementById("pw").innerHTML=itml;
}
divMsg('<%=target%>')
</script>
<%
		else
			response.Write file.Exception
		end if
	end if
end if
set upload = Nothing
%>
<a class="btn2" href="javascript:history.go(-1)">继续上传</a>
</body>
</html>