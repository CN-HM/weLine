<!--#include file="../inc/config.asp"-->
<!DOCtYPE html PUBLIC "-//W3C//Dtd html 4.01 transitional//EN" "http://www.w3c.org/tr/1999/REC-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=SiteName%></title>
<link href="./css/login.css" type="text/css" rel="stylesheet"> 
<meta http-equiv=Content-type content="text/html; charset=gb2312">
<script language="JavaScript">
<!--
function chk(theForm)
{
if (theForm.admin_name.value == "")
{
alert("����������ʺţ�");
theForm.admin_name.focus();
return (false);
}
if (theForm.admin_pass.value == "")
{
alert("������������룡");
theForm.admin_pass.focus();
return (false);
}

return true;
}
//-->
</script>
</head>
<body>
<div class="login">
    <div class="message"><%=SiteName%></div>
    <div id="darkbannerwrap"></div>
    
    <FORM action="check.asp?action=login" method=post onSubmit="return chk(this)">
		<input name="action" value="login" type="hidden">
		<input id=admin_name name=admin_name placeholder="�û���" required="" type="text">
		<hr class="hr15">
		<input id=admin_pass type=password name=admin_pass placeholder="����" required="" >
		<hr class="hr15">
		<input type=text class=pwd id=VerifyCode maxLength=4 name=VerifyCode placeholder="��֤��">
		<hr class="hr15">
		<tr align=right>���ˢ����֤�룺</tr><img src="yz.asp" border='0' onClick="this.src='yz.asp?t='+(new Date().getTime());"  alt='���ˢ����֤��' />
		<hr class="hr15">
		<input type=submit value="��¼" name=Submit  style="width:100%;" >
		<hr class="hr15">
		<input type="Reset" value="ȡ��" name="Reset" style="width:100%;" >
		<hr class="hr20">
		
	</form>
</div>
<div class="copyright"><tr><td align=middle bgColor=#9CBFE5 height=25>2008-2016 &copy; <a target="_blank">ά������ά�޹���ϵͳ</a> All Rights Reserved </td></tr></div>
</body>
</html>