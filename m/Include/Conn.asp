<%
response.buffer=true '���û��崦��

dir="/"'------------��Ŀ¼Ϊ"/",����Ŀ¼ʱ����Ŀ¼�����磺"/domo/"
db=""&dir&"data/zychcms.vip#flagship@#@#!.mdb"'���ݿ�����·��
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.open connstr
Call OpenConn()
If Err Then
err.Clear
Set Conn = Nothing
Response.Write "���ݿ����ӳ����������ݿ������ļ��е����ݿ�������á�"
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
response.write("<script>alert('���󣺽�ֹ��վ���ⲿ�ύ����!.')</script>")
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
upFileSize=20971520 '����ϴ���С��Ĭ����2M
upFileWay=1 	'1:�������Ŀ¼ 2:���´���Ŀ¼ 3:����չ����Ŀ¼
upFiletype="txt,rar,zip,doc,xls,ppt,jpg,jpeg,gif,png,swf,wmv,avi,wma,mp3,mid,flv"'�ϴ���չ��
'��������ThumbnailImg
'�� �ã�����ͼƬ������ͼ
'�� ����ImgUrl ԭͼ��ַ
' ImgWidth ��ͼ�Ŀ�
' Imgheight ��ͼ�ĸ�
' newImgurl ��ͼ�Ĵ�ŵ�ַ
'****************************************************
Sub ThumbnailImg(Imgurl,ImgWidth,Imgheight,newImgurl)
Dim Jpeg ''''//��������
If instr(imgUrl,":\")=0 Then imgUrl = Server.MapPath(imgUrl) 
If instr(newImgurl,":\")=0 Then newImgurl = Server.MapPath(newImgurl) 
Set Jpeg = Server.CreateObject("Persits.jpeg") ''''//�������
Jpeg.Open Imgurl ''''//ԭͼλ��
If ImgWidth<>"" and Imgheight="" Then Imgheight = ImgWidth*Jpeg.OriginalHeight / Jpeg.OriginalWidth 
If Imgheight<>"" and ImgWidth="" Then ImgWidth = Imgheight*Jpeg.OriginalWidth / Jpeg.OriginalHeight 
Jpeg.Width = ImgWidth ''''//��ͼƬ���
Jpeg.Height = Imgheight ''''//��ͼƬ�߶�
Jpeg.Sharpen 1, 130 ''''//�趨��Ч��
Jpeg.Save newImgurl ''''//��������ͼλ�ü�����
Set Jpeg = Nothing ''''//ע��������ͷ���Դ
End Sub
%>