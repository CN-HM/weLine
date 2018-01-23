<%
session.TimeOut=60
if session("admin_name")="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('您还未登录或者超时,请登录！');"&"window.location.href = 'login.asp'"&" </script>"
response.end
end if

sysConfig="维度先生维修管理系统"
%>