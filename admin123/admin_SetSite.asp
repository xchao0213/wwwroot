<!--#include file="../Include/inc.asp" -->
<!--#include file="../Include/Conn.asp" -->
<!--#include file="seeion.asp"-->
<%call chkAdmin(2)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" id="css" href="images/style.css">
<title>附加参数设置</title>
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
</script>
    <script>
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
	function checkEnter(e)
{
var et=e||window.event;
var keycode=et.charCode||et.keyCode;
if(keycode==13)
{
if(window.event)
window.event.returnValue = false;
else
e.preventDefault();//for firefox
}
}
</script>
</head>
<body>
<%
Select Case request.QueryString("Action")
Case "SaveConst"
SaveConstInfo
End Select%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="30" background="images/bg_list.gif"><div style="padding-left:10px; font-weight:bold; color:#FFFFFF">【网站静态功能管理】</div></td>
  </tr>
  <tr>
  <td bgcolor="#FFFFFF">
<form name="ConstForm" method="post" action="?Action=SaveConst">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
<tr>
  <td width="270"  align="right"class="td">系统安装目录：</td>
  <td class="td"><input name="dir" type="text" id="dir" style="width: 280" readonly="readonly" value="<%=dir%>">
    <font color="red">*</font></td>
</tr>

<tr>
  <td class="td" align="right">是否开启系统静态功能：</td>
  <td class="td">
  <label><input name="html" type="radio" onClick="htmlml.style.display=''" value="0" <%if html = 0 then%> checked<%end if%>>启用</label>
  <label><input name="html" type="radio" onClick="htmlml.style.display='none'" value="1" <%if html = 1 then%> checked<%end if%>>不启用 <font color="red">*</font></label>
  <label for="1"><input name="delok" type="checkbox" id="1" value="1" />删除静态页面</label>
</td>
</tr>
  
</table>
<table id="htmlml" style="display:<%If html<>0 Then Response.Write "none"%>"  width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
<tr>
  <td width="270" class="td" align="right">生成静态页面后缀：</td>
  <td class="td"><select style="width: 80" name="HTMLName">
      <option value="html" <%if HTMLName="html" then response.write "selected"%>>html</option>
      <option value="htm" <%if HTMLName="htm" then response.write "selected"%>>htm</option>
      <option value="shtml" <%if HTMLName="shtml" then response.write "selected"%>>shtml</option>
    </select>
    <font color="red">*</font> 如：New.html中“html”即为后缀。可设置如：html、htm、shtml</td>
</tr>
<tr>
  <td align="right" class="td">案例分类页生成前缀：</td>
  <td class="td"><input name="Casefl" type="text" id="Casefl" style="width: 180" value="<%=Casefl%>">
    <font color="red">*</font> 如：Case_1.html中"Case"即为前缀</td>
  </tr>
<tr>
  <td class="td" align="right">案例详细页生成前缀：</td>
  <td class="td"><input name="ShowCase" type="text" id="ShowCase" style="width: 180" value="<%=ShowCase%>">
    <font color="red">*</font> 如：ShowCase_1.html中"ShowCase"即为前缀</td>
  </tr>
<tr>
  <td class="td" align="right">新闻分类页生成前缀：</td>
  <td class="td"><input name="Newsfl" type="text" id="Newsfl" style="width: 180" value="<%=Newsfl%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">新闻详细页生成前缀：</td>
  <td class="td"><input name="ShowNews" type="text" id="ShowNews" style="width: 180" value="<%=ShowNews%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">产品分类页生成前缀：</td>
  <td class="td"><input name="Productsfl" type="text" id="Productsfl" style="width:180" value="<%=Productsfl%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">产品详细页生成前缀：</td>
  <td class="td"><input name="ShowProducts" type="text" id="ShowProducts" style="width: 180" value="<%=ShowProducts%>" />
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">视频分类页生成前缀：</td>
  <td class="td"><input name="Videofl" type="text" id="Videofl" style="width: 180" value="<%=Videofl%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">视频详细页生成前缀：</td>
  <td class="td"><input name="ShowVideo" type="text" id="ShowVideo" style="width:180" value="<%=ShowVideo%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">下载分类页生成前缀：</td>
  <td class="td"><input name="Downloadfl" type="text" id="Downloadfl" style="width:180" value="<%=Downloadfl%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">下载详细页生成前缀：</td>
  <td class="td"><input name="ShowDownload" type="text" id="ShowDownload" style="width:180" value="<%=ShowDownload%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">相册分类页生成前缀：</td>
  <td class="td"><input name="photofl" type="text" id="photofl" style="width: 180" value="<%=photofl%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">相册详细生成前缀：</td>
  <td class="td"><input name="Showphoto" type="text" id="Showphoto" style="width: 180" value="<%=Showphoto%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">服务项目详细生成前缀：</td>
  <td class="td"><input name="Showproject" type="text" id="Showproject" style="width: 180" value="<%=Showproject%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">团队详细生成前缀：</td>
  <td class="td"><input name="Showteam" type="text" id="Showteam" style="width: 180" value="<%=Showteam%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">职位详细生成前缀：</td>
  <td class="td"><input name="Showjob" type="text" id="Showjob" style="width: 180" value="<%=Showjob%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">单页详细生成前缀：</td>
  <td class="td"><input name="about" type="text" id="about" style="width: 180" value="<%=about%>">
    <font color="red">*</font></td>
  </tr>
<tr>
  <td class="td" align="right">所有分页生成分隔：</td>
  <td class="td"><input name="fy" type="text" id="fy" style="width: 180" value="<%=fy%>">
    <font color="red">*</font></td>
</tr>
<td class="td" align="right">分隔符：</td>
  <td class="td"><input name="Separated" type="text" id="Separated" style="width: 80" value="<%=Separated%>">
    <font color="red">*</font> 如：New-1.html中的"-"即为分隔符</td>
  </tr>
</table>
  <div style="padding-left:10px; line-height:30px; font-weight:bold; color:#FFFFFF ; background:url(images/bg_list.gif)">【网站显示设置】</div>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="stable">
<tr>
      <td align="right" width="270" class="td">联系我们公司名称</td>
      <td  class="td"><input name="ditu_title" type="text"  style="width:180" maxlength="8" value="<%=ditu_title%>"/> 最多8个字</td>
    </tr>
  <tr>
      <td align="right"  class="td">百度地图坐标</td>
      <td  class="td"><input name="ditu_Coordinate" type="text"  style="color:#FF0000;width:180" value="<%=ditu_Coordinate%>" />
      [<a href="http://api.map.baidu.com/lbsapi/getpoint/index.html" target="_blank"  style="color:#003399; ">获取地图坐标</a>] * 注：获取的106.510742,29.542198 坐标</td>
    </tr>
    <tr>
      <td align="right" class="td">联系我们地图名称</td>
      <td class="td"><input name="ditu_mc" type="text" style="width:180" value="<%=ditu_mc%>" /></td>
    </tr>
    <tr>
      <td  align="right" class="td">联系我们地图说明</td>
      <td  class="td"><textarea name="ditu_sm" cols="50"  onkeydown="checkEnter(event)"><%=ditu_sm%></textarea></td>
    </tr>
    <tr>
      <td  align="right" class="td">官方微博地址</td>
      <td  class="td"><input name="zych_weibo" type="text" value="<%=zych_weibo%>" size="50" /></td>
    </tr>
  <tr>
    <td class="td" align="right"><font color="red">*</font>首页单页内容调用ID</td>
    <td class="td"><input name="sy_dy_id" type="text" id="sy_dy_id" style="width: 80" value="<%=sy_dy_id%>"></td>
  </tr>
  <tr>
    <td class="td" align="right"><font color="red">*</font>首页单页内容调用ID字数</td>
    <td class="td"><input name="sy_dy_zs" type="text" id="sy_dy_zs" style="width: 80" value="<%=sy_dy_zs%>"></td>
  </tr>

  <tr>
    <td class="td" align="right"><font color="red">*</font>首页大图滚动图文</td>
    <td class="td"><select name="Scroll">
        <option value="1" <%if Scroll=1 then%>selected<%end if%>>案例推荐滚动</option>
        <option value="2" <%if Scroll=2 then%>selected<%end if%>>产品推荐滚动</option>
      </select> 标题 <input name="in_title" type="text" style="width:100px;text-align:center" maxlength="8" value="<%=in_title%>" /> 
      行数<input name="in_hs" type="text"  style="width:80px;text-align:center" maxlength="8" value="<%=in_hs%>" /></td>
  </tr>
  <tr>
    <td class="td" align="right">文章频道列表风格</td>
    <td class="td"><select name="newslb">
        <option value="1" <%if newslb=1 then%>selected<%end if%>>文字列表</option>
        <option value="2" <%if newslb=2 then%>selected<%end if%>>图文列表</option>
      </select></td>
  </tr> 
  
  <tr bgcolor="#ffffff">
        <td height="25" align="right" class="td">首页视频播放设置</td>
        <td class="td"><input type="radio" name="vid_pay" value="1" <%if vid_pay=1 then%>checked<%end if%> />自动
  <input type="radio" name="vid_pay" value="0" <%if vid_pay=0 then%>checked<%end if%> />点击
  <span style="padding-left:20px;">视频显示网址：<input name="vid_wz" type="text"  value="<%=vid_wz%>" /></span></td>
      </tr>

  <tr>
      <td height="25" align="right" class="td">是否开放注册</td>
      <td class="td"><input type="radio" name="userzc" value="0" <%if userzc=0 then%>checked<%end if%> />开放
        <input type="radio" name="userzc" value="1" <%if userzc=1 then%>checked<%end if%> />关闭</td>
    </tr>
   <tr>
      <td height="25" align="right" class="td">是否开启手机自动访问手机版</td>
      <td class="td"><input type="radio" name="wap" value="0" <%if wap=0 then%>checked<%end if%> />开启
        <input type="radio" name="wap" value="1" <%if wap=1 then%>checked<%end if%> />关闭</td>
    </tr>
    <tr>
      <td height="25" align="right" class="td">会员注册是否需要审核</td>
      <td  class="td"><input type="radio" name="usersh" value="1" <%if usersh=1 then%>checked<%end if%> />无需
          <input type="radio" name="usersh" value="0" <%if usersh=0 then%>checked<%end if%> />需要</td>
    </tr>
  <tr>
    <td class="td" align="right">频道页显示友情链接</td>
    <td class="td"><input name="link" type="radio" value="0" <%if link = 0 then%> checked<%end if%>>显示
      <input name="link" type="radio" value="1" <%if link = 1 then%> checked<%end if%>>不显示</td>
  </tr>
  
  <tr>
    <td class="td" align="right">后台登陆是否开启验证码</td>
    <td class="td"><input name="code" type="radio" value="0" <%if code = 0 then%> checked<%end if%>>开启
    <input name="code" type="radio" value="1" <%if code = 1 then%> checked<%end if%>>关闭</td>
  </tr>
  <tr >
      <td align="right" class="td">网站客服功能是否开启</td>
      <td class="td"><input name="qqkf" type="radio" value="0" <%if qqkf = 0 then%> checked<%end if%>>开启
          <input name="qqkf" type="radio" value="1" <%if qqkf = 1 then%> checked<%end if%>>关闭</td>
    </tr>
      
  <tr>
    <td class="td" align="right">是否开启二维码显示</td>
    <td class="td"><input name="zych_rwmon" type="radio" value="0"  onClick="Date3.style.display=''" <%if zych_rwmon= 0 then%> checked<%end if%>>显示
                   <input name="zych_rwmon" type="radio" value="1" onClick="Date3.style.display='none'" <%if zych_rwmon= 1 then%> checked<%end if%>>不显示
                   
  <table width="100" border="0" cellspacing="0" cellpadding="0" id="Date3" style="display:<%If zych_rwmon<>0 Then Response.Write "none"%>">
        <tr>
            <td><input name="zych_rwmurl" type="text"  value="<%=zych_rwmurl%>" size="40" /></td>
            <td><iframe src="up.asp?formname=ConstForm&inputname=zych_rwmurl&uploadstyle=file" frameborder="0" scrolling="no" width="300" height="25"></iframe></td>
          </tr>
      </table></td>
  </tr>
      <tr>
    <td class="td" align="right">是否开启图片水印功能</td>
    <td class="td"><input name="Watermark" type="radio" value="0" onClick="Date4.style.display=''" <%if Watermark= 0 then%> checked<%end if%>>开启
                   <input name="Watermark" type="radio" value="1" onClick="Date4.style.display='none'" <%if Watermark= 1 then%> checked<%end if%>>关闭&nbsp;&nbsp; &nbsp;&nbsp;
      <span id="Date4" style="display:<%If Watermark<>0 Then Response.Write "none"%>"> 
      图片大小 <input type="text" name="Watermark_w" value="<%=Watermark_w%>" size="5" /> ×
      <input type="text" name="Watermark_h" value="<%=Watermark_h%>" size="5" />
      <table width="42%" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td width="21%"><input type="text" name="Watermark_url" value="<%=Watermark_url%>" size="40" /></td>
  <td width="79%"><iframe src="up.asp?formname=ConstForm&inputname=Watermark_url&uploadstyle=file" frameborder="0" scrolling="no" width="300" height="25"></iframe></td>
</tr>
</table></span>	  </td>
  </tr>
  <tr>
  <td width="14%" height="25" align="right" class="td">网站背景风格</td>
  <td width="86%" class="td">
  <input name="body_no" type="radio" value="0"  onClick="data5.style.display='none'" <%if body_no= 0 then%> checked<%end if%>>默认
  <input name="body_no" type="radio" value="1" onClick="data5.style.display=''" <%if body_no= 1 then%> checked<%end if%>>自定
    <table width="452" border="0" cellspacing="0" cellpadding="0" id="data5" style="display:<%If body_no<>1 Then Response.Write "none"%>">
      <tr>
        <td width="245"><input name="body_bj" id="body_bj" type="text" value="<%=body_bj%>" size="35" /></td>
        <td width="95"><input type="button" id="filemanager" value="选取已上传图片" class="btn" style="height:20px;" /></td>
        <td width="112"><iframe src="up.asp?formname=ConstForm&amp;inputname=body_bj&amp;uploadstyle=file" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
      </tr>
      <tr>
        <td width="245"><input name="body_color" id="colorPicker" type="text" value="<%=body_color%>" size="35" /><script type="text/javascript" src="js/Deepteach_colorPicker.js"></script></td>
        <td colspan="2"> 背景色 墨绿色：#000D15 淡灰色:#F6F6F6</td>
        </tr>
  </table></td>
  </tr>
    <tr>
   <td background="images/bg_list.gif"><div style="padding-left:10px; font-weight:bold; color:#FFFFFF">【网站运营设置】</div></td>
   <td background="images/bg_list.gif"></td>
</tr>
<tr>
<tr >
<td width="18%" height="13" align="right" class="td">是否开启在线(支付)购物功能：</td>
<td width="82%"  class="td"><input name="pay" type="radio" value="0" <%if pay= 0 then%> checked<%end if%>>开启
    <input name="pay" type="radio" value="1" <%if pay=1 then%> checked<%end if%>>关闭</td>
</tr>
<tr>
<td align="right" class="td">支付宝接口类型</td>
<td class="td"><select name="alipayLX">
		<option value="danbao" <%if alipayLX="danbao" then%>selected<%end if%>>担保交易</option>
        <option value="Guarantee" <%if alipayLX="Guarantee" then%>selected<%end if%>>双功能收款</option>
        <!--<option value="Immediate" <%if alipayLX="Immediate" then%>selected<%end if%>>即时到账收款</option>-->
      </select> <a href="tencent://message/?uin=495573637&Site=ziliaok8&Menu=yes" target="_blank">即时到账接口必须为企业支付宝帐号，需单独定制</a></td>
</tr>
<tr>
<td align="right" class="td">收款支付宝帐号</td>
<td class="td"><input name="alipayID" type="text" style="width:180" value="<%=alipayID%>" /></td>
</tr>
<tr>
<td align="right" class="td">合作者身份(PID)</td>
<td class="td"><input name="alipayPID" type="password" style="width:180" value="<%=alipayPID%>" /> 
仅支持"双功能收款"与“担保交易收款” 功能能接口 <a href="https://b.alipay.com/order/productIndex.htm" style="color:#F00" target="_blank">没有点击申请</a></td>
</tr>
<tr>
<td align="right" class="td">安全校验码(Key)</td>
<td class="td"><input name="alipayKey" type="password" style="width:260" value="<%=alipayKey%>" /></td>
</tr>
<tr>
<td align="right" class="td">快递查询接口Key</td>
<td class="td"><input name="kdappKey" type="text" style="width:260" value="<%=kdappKey%>" /> 没有接口码，在这里申请：<a href="http://www.kuaidi.com/openapi.html">http://www.kuaidi.com/openapi.html</a></td>
</tr>

  <tr>
    <td class="td" align="right">&nbsp;</td>
    <td class="td"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn" value="保存附加参数设置"></td>
  </tr>
</table>
</form>
</td>
  </tr>
</table>
<%
Function SaveConstInfo()
if Trim(request("in_hs"))="" then
response.Write("<script language=javascript>alert('行数不能为空!');history.go(-1)</script>") 
response.end 
end if
    Set ASM = server.CreateObject("Adod" & "b.St" & "ream")
    ASM.Type = 2
    ASM.mode = 3
    ASM.charset = "GB2312"
    ASM.Open
    ASM.WriteText "<" & "%" & vbCrLf
    ASM.WriteText "Const html = " & ReplaceConstChar(Trim(request("html"))) & "  '------开启网站静态" & vbCrLf
	ASM.WriteText "Const HTMLdir = " & Chr(34) & ReplaceConstChar(Trim(request("HTMLdir"))) & Chr(34) & "   '------网站静态目录" & vbCrLf
    ASM.WriteText "Const sy_dy_id = " & ReplaceConstChar(Trim(request("sy_dy_id"))) & "" & vbCrLf
    ASM.WriteText "Const sy_dy_zs = " & ReplaceConstChar(Trim(request("sy_dy_zs"))) & "" & vbCrLf
    ASM.WriteText "Const HTMLName = " & Chr(34) & ReplaceConstChar(Trim(request("HTMLName"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Casefl = " & Chr(34) & ReplaceConstChar(Trim(request("Casefl"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const ShowCase = " & Chr(34) & ReplaceConstChar(Trim(request("ShowCase"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Newsfl = " & Chr(34) & ReplaceConstChar(Trim(request("Newsfl"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ShowNews = " & Chr(34) & ReplaceConstChar(Trim(request("ShowNews"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Productsfl = " & Chr(34) & ReplaceConstChar(Trim(request("Productsfl"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ShowProducts = " & Chr(34) & ReplaceConstChar(Trim(request("ShowProducts"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Videofl = " & Chr(34) & ReplaceConstChar(Trim(request("Videofl"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ShowVideo = " & Chr(34) & ReplaceConstChar(Trim(request("ShowVideo"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Downloadfl = " & Chr(34) & ReplaceConstChar(Trim(request("Downloadfl"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ShowDownload = " & Chr(34) & ReplaceConstChar(Trim(request("ShowDownload"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Planningfl = " & Chr(34) & ReplaceConstChar(Trim(request("Planningfl"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ShowPlanning = " & Chr(34) & ReplaceConstChar(Trim(request("ShowPlanning"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const photofl = " & Chr(34) & ReplaceConstChar(Trim(request("photofl"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Showphoto = " & Chr(34) & ReplaceConstChar(Trim(request("Showphoto"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Showteam = " & Chr(34) & ReplaceConstChar(Trim(request("Showteam"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Showjob = " & Chr(34) & ReplaceConstChar(Trim(request("Showjob"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const Showproject = " & Chr(34) & ReplaceConstChar(Trim(request("Showproject"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const about = " & Chr(34) & ReplaceConstChar(Trim(request("about"))) & Chr(34) & "" & vbCrLf'-------------------------------------------
    ASM.WriteText "Const fy = " & Chr(34) & ReplaceConstChar(Trim(request("fy"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const Separated = " & Chr(34) & ReplaceConstChar(Trim(request("Separated"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const body_no = "&ReplaceConstChar(Trim(request("body_no")))&"" & vbCrLf
	ASM.WriteText "Const body_bj = "&Chr(34)&ReplaceConstChar(Trim(request("body_bj")))& Chr(34) & "" & vbCrLf
	ASM.WriteText "Const body_color= " & Chr(34) & ReplaceConstChar(Trim(request("body_color")))& Chr(34) & "" & vbCrLf
	ASM.WriteText "Const link = " & ReplaceConstChar(Trim(request("link"))) & "" & vbCrLf
	ASM.WriteText "Const ditu_title= " & Chr(34) & ReplaceConstChar(Trim(request("ditu_title")))& Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ditu_Coordinate= " & Chr(34) & ReplaceConstChar(Trim(request("ditu_Coordinate")))& Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ditu_mc = " & Chr(34) & ReplaceConstChar(Trim(request("ditu_mc"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ditu_sm = " & Chr(34) & ReplaceConstChar(Trim(request("ditu_sm"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const zych_weibo = " & Chr(34) & ReplaceConstChar(Trim(request("zych_weibo"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const code = " & ReplaceConstChar(Trim(request("code"))) &"" & vbCrLf
	ASM.WriteText "Const userzc = " & ReplaceConstChar(Trim(request("userzc"))) &"" & vbCrLf
	ASM.WriteText "Const usersh = " & ReplaceConstChar(Trim(request("usersh"))) &"" & vbCrLf
	ASM.WriteText "Const qqkf = " & ReplaceConstChar(Trim(request("qqkf"))) &"" & vbCrLf
	ASM.WriteText "Const Scroll = " & ReplaceConstChar(Trim(request("Scroll"))) &"" & vbCrLf
	ASM.WriteText "Const in_title= "&Chr(34)&ReplaceConstChar(Trim(request("in_title")))&Chr(34)&vbCrLf
	ASM.WriteText "Const in_hs= "&Chr(34)&ReplaceConstChar(Trim(request("in_hs")))&Chr(34)&vbCrLf
	ASM.WriteText "Const newslb = "&Chr(34)& ReplaceConstChar(Trim(request("newslb")))&Chr(34)& vbCrLf
	ASM.WriteText "Const vid_pay = "&Chr(34)& ReplaceConstChar(Trim(request("vid_pay")))&Chr(34)&vbCrLf
	ASM.WriteText "Const vid_wz= "&Chr(34)& ReplaceConstChar(Trim(request("vid_wz")))&Chr(34)&vbCrLf 
	ASM.WriteText "Const wap = " &Chr(34)&ReplaceConstChar(Trim(request("wap")))&Chr(34)&vbCrLf 

	ASM.WriteText "Const zych_rwmon = " & Chr(34) & ReplaceConstChar(Trim(request("zych_rwmon"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const zych_inpro = " & Chr(34) & ReplaceConstChar(Trim(request("zych_inpro"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const zych_rwmurl = " & Chr(34) & ReplaceConstChar(Trim(request("zych_rwmurl"))) & Chr(34) & "" & vbCrLf 
	ASM.WriteText "Const Watermark = " & ReplaceConstChar(Trim(request("Watermark"))) &"" & vbCrLf 
    ASM.WriteText "Const Watermark_url = " & Chr(34)& ReplaceConstChar(Trim(request("Watermark_url")))& Chr(34) &"" & vbCrLf
	ASM.WriteText "Const Watermark_w = " & Chr(34)& ReplaceConstChar(Trim(request("Watermark_w")))& Chr(34) &"" & vbCrLf
	ASM.WriteText "Const Watermark_h = " & Chr(34)& ReplaceConstChar(Trim(request("Watermark_h")))& Chr(34) &"" & vbCrLf
	ASM.WriteText "Const pay= " & ReplaceConstChar(Trim(request("pay"))) &"" & vbCrLf 
	ASM.WriteText "Const alipayLX = " & Chr(34) & ReplaceConstChar(Trim(request("alipayLX"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const alipayID = " & Chr(34) & ReplaceConstChar(Trim(request("alipayID"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const alipayPID = " & Chr(34) & ReplaceConstChar(Trim(request("alipayPID"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const alipayKey = " & Chr(34) & ReplaceConstChar(Trim(request("alipayKey"))) & Chr(34) & "" & vbCrLf
	ASM.WriteText "Const kdappKey = " & Chr(34) & ReplaceConstChar(Trim(request("kdappKey"))) & Chr(34) & "" & vbCrLf

    ASM.WriteText "%" & ">"
    ASM.SaveToFile Server.mappath("../Include/inc.asp"), 2
    ASM.flush
    ASM.Close
    Set ASM = Nothing
    If ReplaceConstChar(Trim(request("HTML"))) = 1 and request.form("delok")= 1 Then Call FsoDelHtml(ReplaceConstChar(Trim(request("HTMLName"))))
    response.Write "<script language=""javascript"">alert('系统附加参数生成成功');location.href='admin_SetSite.asp';</script>"
End Function

Sub FsoDelHtml(HTMLName)
    Dim Fso, FsoOut, File
    Set Fso = Server.CreateObject("Scripting.FileSystemObject")
    Set FsoOut = Fso.GetFolder(Server.Mappath(""&Dir&""))
	For Each File In FsoOut.Files
        If Right(File.Name, Len(File.Name) - InStrRev(File.Name, ".")) = HTMLName And HTMLName <> "asp" Then
            Response.Write "<span style=""color:red; padding-left: 18px"">" & File.Name & "</span> 成功删除…<br />"
            Fso.DeleteFile File.Path, True
        End If
    Next
	htmlml="About,Case,Contact,Download,Job,News,Photo,Products,Service,Team,Video" 
	htmlml=split(htmlml,",") 
	for i=0 to ubound(htmlml) 
		Set FsoOut = Fso.GetFolder(Server.Mappath(""&Dir&htmlml(i)&""))
		For Each File In FsoOut.Files
			If Right(File.Name, Len(File.Name) - InStrRev(File.Name, ".")) = HTMLName And HTMLName <> "asp" Then
				Response.Write "<span style=""color:red; padding-left: 18px"">"&htmlml(i)&"/"&File.Name&"</span> 成功删除…<br />"
				Fso.DeleteFile File.Path, True
			End If
		Next
	next 
	Set FsoOut = Nothing
    Set Fso = Nothing
End Sub

Function ReplaceConstChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceConstChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceConstChar = tempChar
End Function
%>
