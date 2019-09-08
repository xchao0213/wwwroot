<%if session("username")<>"" then
else
session("Tourl")=Request.ServerVariables("HTTP_REFERER")
response.write "<script>;window.location.href='/User/login.asp';</script>"
response.end
end if
Function DelHtml(Str1)
  Dim regEx
  Set regEx = New RegExp
  regEx.Pattern = "(<[^>]*?>)"
  regEx.Global = True
  regEx.IgnoreCase = True
  DelHtml = replace(regEx.Replace(""&str1,""),"&nbsp;","")
End Function

Function InterceptString(txt,length)
    txt=trim(txt)
    x = len(txt)
    y = 0
    if x >= 1 then
        for ii = 1 to x
            if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '如果是汉字
                y = y + 2
            else
                y = y + 1
            end if
            if y >= length then
                txt = left(trim(txt),ii) '字符串限长
                exit for
            end if
        next
        InterceptString = txt
    else
        InterceptString = ""
    end if
End Function
%>