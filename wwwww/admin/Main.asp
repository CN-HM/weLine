<!--#include file="inc/right.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 transitional//EN" "http://www.w3.org/tr/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sysConfig%></title>
<link href="images/main.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td  bgcolor="#FFFFFF"><br>
	<table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
	  <tr align="center" bgcolor="#F2FDFF">
        <td class="optiontitle" colspan="4">系统检测</td>	
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>后台操作管理员：</td>
        <td colspan="3"><%=session("admin_name")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="100">服务器名：</td>
        <td width="250"><%=Request.ServerVariables("SERVER_NAME")%></td>
        <td width="20%">服务器IP：</td>
        <td><%=Request.ServerVariables("LOCAL_ADDR")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>服务器端口：</td>
        <td><%=Request.ServerVariables("SERVER_PORT")%></td>
        <td>服务器时间：</td>
        <td><%=now%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>本文件绝对路径：</td>
        <td colspan="3"><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>客户端IP：</td>
        <td><%=Request.ServerVariables("REMOTE_ADDR")%></td>
        <td>端口：</td>
        <td><%=Request.ServerVariables("REMOTE_PORT")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>浏览器版本：</td>
        <td colspan="3"><%=Request.ServerVariables("Http_User_Agent")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>脚本解释引擎：</td>
        <td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
        <td>Jmail邮箱组件支持：</td>
        <td><%If IsObjInstalled("JMail.Message") Then%>
          Jmail4.3邮箱组件支持
            <%elseIf IsObjInstalled("JMail.SMTPMail") then%>
            Jmail4.2组件支持
            <%else%>
            不支持Jmail组件
          <%end if%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td bgcolor="#FFFFFF">FSO组件支持：</td>
        <td>
<%If IsObjInstalled("Scripting.FileSystemObject") then%>FSO支持
<%else%>不支持FSO组件
<%end if%>
<%Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = true
Set xTestObj = Nothing
Err = 0
End Function
%></td>
        <td>CDONTS邮箱组件支持：</td>
        <td><%If IsObjInstalled("CDONTS.NewMail") then%>CDONTS邮箱组件支持
            <%else%>不支持CDONTS邮箱组件<%end if%></td>
      </tr>
    </table>
	<p>
    <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
	  <tr align="center" bgcolor="#F2FDFF">
        <td class="optiontitle" colspan="2">系统信息</td>	
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="100"> 系统名称：</td>
        <td><%=sysConfig%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td > 程序版本：</td>
        <td>V2.4 </td>		
      </tr>
	  
    </table>
	<p>
	<table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
	  <tr align="center" bgcolor="#F2FDFF">
        <td class="optiontitle" colspan="2">系统维护员信息日志</td>	
      </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>维护时间</td>
		<td>维护内容</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2016年11月29日</td>
		<td>1、增加了维修员查询 <br> 2、修正了一个JS：checkbox全选错误 <br> 3、相关布局调整 <br>4、版本更新至2.2</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2016年12月1日</td>
		<td>1、更新了微信访问页面，移动端登记已经上线 <br> 2、发现一个FSO错误，暂无解决方案 <br>3、版本更新至2.3</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2016年12月7日</td>
		<td>1、查询页面直接可以看到维修人员信息<br>2、版本更新至2.4</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2017年1月8日</td>
		<td>1、更新前台界面<br>2、优化了前端访问速度及后台细节<>3、版本更新至3.1</td>
	  </tr>

    </table>
	
	</td>
  </tr>
</table>
</body>
</html>