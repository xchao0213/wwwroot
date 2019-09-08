<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(1)
set config=server.createobject("adodb.recordset") 
exec="select * from  config  " 
config.open exec,conn,1,1
WebURL = "http://"&request.servervariables("server_name")&"" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站基本设置</title>
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<link rel="stylesheet" href="../zycheditor/themes/default/default.css" />
<link rel="stylesheet" href="../zycheditor/plugins/code/prettify.css" />
<script charset="utf-8" src="../zycheditor/kindeditor.js"></script>
<script charset="utf-8" src="../zycheditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zycheditor/plugins/code/prettify.js"></script>
<script>
KindEditor.ready(function(K) {
    var editor1 = K.create('textarea[name="body"]', {
        cssPath : '../zycheditor/plugins/code/prettify.css',
        uploadJson : '../zycheditor/asp/upload_json.asp',
        fileManagerJson : '../zycheditor/asp/file_manager_json.asp',
        allowFileManager : true,
        afterCreate : function() {
            var self = this;
            K.ctrl(document, 13, function() {
                self.sync();
                K('form[name=add]')[0].submit();
            });
            K.ctrl(self.edit.doc, 13, function() {
                self.sync();
                K('form[name=add]')[0].submit();
            });
        }
    });
    prettyPrint();
});

KindEditor.ready(function(K) {
        var editor = K.editor({
            fileManagerJson : '../zycheditor/asp/file_manager_json.asp'
        });
        K('#filemanager').click(function() {
            editor.loadPlugin('filemanager', function() {
                editor.plugin.filemanagerDialog({
                    viewType : 'VIEW',
                    dirName : 'image',
                    clickFn : function(url, title) {
                        K('#body_bj').val(url);
                        editor.hideDialog();
                    }
                });
            });
        });
    });
</script>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div  style="padding-left:10px; font-weight:bold; color:#FFFFFF">网站基本设置</div></td>
  </tr>
  <tr>
    <td>
<form  name="add" method="post" action="admin_xtsz.asp?wzsz=ok">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF" class="stable">
      <tr>
        <td width="13%" height="25" align="right" class="td">网站名称</td>
        <td width="87%"  class="td"><input name="title" type="text" value="<%=config("title")%>" size="40" />
        <input name="id" type="hidden" value="<%=config("id")%>" size="50" /></td>
      </tr>
      <tr>
        <td width="13%" height="7" align="right" class="td">网站域名</td>
        <td class="td"><input name="url" type="text"  value="<%=config("url")%>" size="40" />
		<input type=button name=button class="btn2" value="设为当前域名" onclick="document.add.url.value='<%=WebURL%>'">你当前的域名为：<%=WebURL%></td>
      </tr>
      <tr>
        <td width="13%" height="6" align="right" class="td">网站LOGO</td>
        <td class="td">
          <table width="200" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><input type="text" value="<%=config("logo")%>" name="logo" size="40" /></td>
              <td><iframe src="up.asp?formname=add&inputname=logo&uploadstyle=Logo" frameborder="0" scrolling="no" width="280" height="25"></iframe></td>
            </tr>
          </table></td>
        </tr>
      <tr>
        <td width="13%" height="3" align="right" class="td">LOGO尺寸</td>
        <td class="td"><input name="logoiwidth" type="text"  value="<%=config("logoiwidth")%>" size="10" />×
          <input name="logoheight" type="text"  value="<%=config("logoheight")%>" size="10" /></td>
      </tr>
	    <tr>
        <td height="25" align="right" class="td">网页关键字</td>
        <td class="td"><textarea name="keywords" cols="40" rows="3"><%=config("keywords")%></textarea></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">网站描述</td>
        <td class="td"><textarea name="description" cols="40" rows="3"><%=config("description")%></textarea></td>
      </tr>
	  
      <tr>
        <td height="25" align="right" class="td"><p>底部版权信息<br>[支持HTML]</p>          </td>
        <td class="td"><textarea name="copyright" cols="40" rows="3"><%=config("copyright")%></textarea></td>
      </tr>
	  <tr>
        <td width="13%" height="25" align="right" class="td">联 系 人</td>
        <td class="td"><input name="name" type="text" value="<%=config("name")%>" size="40" /></td>
      </tr>
	   <tr>
        <td height="25" align="right" class="td">联系手机</td>
        <td class="td"><input name="Phone" type="text" value="<%=config("Phone")%>" size="40" /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">公司邮箱</td>
        <td class="td"><input name="mail" type="text" value="<%=config("mail")%>"  size="40" /></td>
      </tr>

      <tr>
        <td height="25" align="right" class="td">公司电话</td>
        <td class="td"><input name="tel" type="text" value="<%=config("tel")%>" size="40" /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">公司传真</td>
        <td class="td"><input name="fax" type="text"  value="<%=config("fax")%>" size="40" /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">邮政编码</td>
        <td class="td"><input name="Postal" type="text"  value="<%=config("Postal")%>" size="40" /></td>
      </tr>
           <tr>
        <td height="25" align="right" class="td">公司地址</td>
        <td class="td"><input name="dz" type="text" id="textfield8" value="<%=config("dz")%>" size="40" /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">ICP备案编号</td>
        <td class="td"><input name="beian" type="text" value="<%=config("beian")%>" /></td>
      </tr>
      <tr>
        <td height="25" align="right" class="td">网站访问量</td>
        <td class="td"><input name="js" type="text"  value="<%=config("js")%>" /></td>
      </tr>

      <tr>
        <td height="25" align="right" class="td">是否开放注册</td>
        <td class="td"><input type="radio" name="userzc" value="0" <%if config("userzc")=0 then%>checked<%end if%> />开放注册　
  <input type="radio" name="userzc" value="1" <%if config("userzc")=1 then%>checked<%end if%> />关闭注册　</td>
      </tr>
	  
       <tr>
        <td height="25" align="right" class="td">会员注册审核</td>
        <td  class="td"><input type="radio" name="usersh" value="1" <%if config("usersh")=1 then%>checked<%end if%> />无需审核　
  <input type="radio" name="usersh" value="0" <%if config("usersh")=0 then%>checked<%end if%> />需要审核</td>
      </tr>
       
      <tr>
        <td height="25" align="right" class="td">留言是否审核</td>
        <td class="td"><input type="radio" name="booksh" value="1" <%if config("booksh")=1 then%>checked<%end if%> />无需审核　
  <input type="radio" name="booksh" value="0" <%if config("booksh")=0 then%>checked<%end if%> />需要审核</td>
      </tr>

<tr bgcolor="#FFFFFFF">
        <td height="25" align="right" class="td" >是否关闭网站</td>
        <td class="td"><input type="radio" name="on" value="0" <%if config("on")=0 then%>checked<%end if%> />开放网站　
  <input type="radio" name="on" value="1" <%if config("on")=1 then%>checked<%end if%> />关闭网站</td>
      </tr>
	  <tr>
        <td height="25" align="right" class="td">关闭说明</td>
        <td class="td"><textarea name="off_sm" cols="40" rows="2"><%=config("off_sm")%></textarea></td>
    </tr>
	  <tr>
        <td width="13%" height="25" align="right" class="td">网站统计代码</td>
        <td width="87%" class="td"><textarea name="zztj" cols="50" rows="3"><%=config("zztj")%></textarea> * 注:不要用“站长统计”</td>
	  </tr>	  
	  <tr>
        <td height="25" align="right" class="td">LOGO右边信息[支持HTML]</td>
        <td class="td"><textarea name="logoright" cols="50" rows="3"><%=config("logoright")%></textarea></td>
	  </tr>		
	  <tr>
        <td height="25" align="right" class="td">网站广告语[支持HTML]</td>
        <td class="td"><textarea name="foot_left" cols="50" rows="3"><%=config("foot_left")%></textarea></td>
	  </tr>		
   
      <tr>
        <td height="47" class="td">&nbsp;</td>
        <td class="td"><label><input type="submit" name="button" id="button" value="保存配置" class="btn" /></label></td>
      </tr>
    </table>
    </form>	
    </td>
  </tr>
</table>
</body>
</html>
<% 
if Request.QueryString("wzsz")="ok" then
id=request("id")  
sql="select * from config where id="&id 
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
zztj=request("zztj")
if Instr(zztj,"cnzz.com")>0 then  
Response.Write "<script language=javascript>alert('请不要用站长统计，没看到右边有说明吗？');history.go(-1)</script>"
response.end 
end if 
title=request("title")
if len(title)>80 then  
Response.Write "<script language=javascript>alert('网站优化建议，网站标题不搬不要超过80个字？');history.go(-1)</script>"
response.end 
end if
rs("title")=request.form("title")
rs("url")=request.form("url") 
rs("logo")=request.form("logo") 
rs("logoiwidth")=request.form("logoiwidth") 
rs("logoheight")=request.form("logoheight") 
rs("css")=request.form("css")
rs("name")=request.form("name")
rs("Phone")=request.form("Phone")
rs("mail")=request.form("mail")
rs("tel")=request.form("tel")
rs("fax")=request.form("fax")
rs("Postal")=request.form("Postal")
rs("dz")=request.form("dz")
rs("copyright")=request.form("copyright")  
rs("beian")=request.form("beian") 
rs("js")=request.form("js") 
rs("keywords")=request.form("keywords")  
rs("description")=request.form("description") 
rs("userzc")=request.form("userzc") 
rs("on")=request.form("on") 
rs("off_sm")=request.form("off_sm") 
rs("usersh")=request.form("usersh") 
rs("booksh")=request.form("booksh") '----------以下开始
rs("logoright")=request.form("logoright")
rs("zztj")=request.form("zztj")
rs("daohangnr")=request.form("daohangnr")
rs("foot_left")=request.form("foot_left") 
rs.update 
rs.close 
response.write "<script>alert('网站设置成功!');window.location.href='admin_xtsz.asp';</script>"  
end if
%> 
