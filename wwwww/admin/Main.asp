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
        <td class="optiontitle" colspan="4">ϵͳ���</td>	
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>��̨��������Ա��</td>
        <td colspan="3"><%=session("admin_name")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="100">����������</td>
        <td width="250"><%=Request.ServerVariables("SERVER_NAME")%></td>
        <td width="20%">������IP��</td>
        <td><%=Request.ServerVariables("LOCAL_ADDR")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>�������˿ڣ�</td>
        <td><%=Request.ServerVariables("SERVER_PORT")%></td>
        <td>������ʱ�䣺</td>
        <td><%=now%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>���ļ�����·����</td>
        <td colspan="3"><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>�ͻ���IP��</td>
        <td><%=Request.ServerVariables("REMOTE_ADDR")%></td>
        <td>�˿ڣ�</td>
        <td><%=Request.ServerVariables("REMOTE_PORT")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>������汾��</td>
        <td colspan="3"><%=Request.ServerVariables("Http_User_Agent")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td>�ű��������棺</td>
        <td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
        <td>Jmail�������֧�֣�</td>
        <td><%If IsObjInstalled("JMail.Message") Then%>
          Jmail4.3�������֧��
            <%elseIf IsObjInstalled("JMail.SMTPMail") then%>
            Jmail4.2���֧��
            <%else%>
            ��֧��Jmail���
          <%end if%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td bgcolor="#FFFFFF">FSO���֧�֣�</td>
        <td>
<%If IsObjInstalled("Scripting.FileSystemObject") then%>FSO֧��
<%else%>��֧��FSO���
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
        <td>CDONTS�������֧�֣�</td>
        <td><%If IsObjInstalled("CDONTS.NewMail") then%>CDONTS�������֧��
            <%else%>��֧��CDONTS�������<%end if%></td>
      </tr>
    </table>
	<p>
    <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
	  <tr align="center" bgcolor="#F2FDFF">
        <td class="optiontitle" colspan="2">ϵͳ��Ϣ</td>	
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="100"> ϵͳ���ƣ�</td>
        <td><%=sysConfig%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td > ����汾��</td>
        <td>V2.4 </td>		
      </tr>
	  
    </table>
	<p>
	<table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
	  <tr align="center" bgcolor="#F2FDFF">
        <td class="optiontitle" colspan="2">ϵͳά��Ա��Ϣ��־</td>	
      </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>ά��ʱ��</td>
		<td>ά������</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2016��11��29��</td>
		<td>1��������ά��Ա��ѯ <br> 2��������һ��JS��checkboxȫѡ���� <br> 3����ز��ֵ��� <br>4���汾������2.2</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2016��12��1��</td>
		<td>1��������΢�ŷ���ҳ�棬�ƶ��˵Ǽ��Ѿ����� <br> 2������һ��FSO�������޽������ <br>3���汾������2.3</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2016��12��7��</td>
		<td>1����ѯҳ��ֱ�ӿ��Կ���ά����Ա��Ϣ<br>2���汾������2.4</td>
	  </tr>
	  <tr olspan="2" bgcolor="#FFFFFF">
		<td>2017��1��8��</td>
		<td>1������ǰ̨����<br>2���Ż���ǰ�˷����ٶȼ���̨ϸ��<>3���汾������3.1</td>
	  </tr>

    </table>
	
	</td>
  </tr>
</table>
</body>
</html>