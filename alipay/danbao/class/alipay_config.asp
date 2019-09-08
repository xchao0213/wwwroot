<!--#include file="../Include/config.asp" -->
<%
partner         = alipayPID

'安全检验码，以数字和字母组成的32位字符
key   			= alipayKey

'字符编码格式 目前支持 gbk 或 utf-8
input_charset	= "gbk"

'签名方式 不需修改
sign_type       = "MD5"

%>