<!--#include file="inc/right.asp"--> 
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/md5.asp"-->
<script language=Javascript>
<!--
function DoEmpty(params)
{
if (confirm("���Ҫɾ��������¼��ɾ����˼�¼����������ݶ�����ɾ�������޷��ָ���"))
window.location = params ;
}
//-->
</script>
<%
if Request("wor")="del" then
sql="delete from Admin where id="&Request("id")
Conn.execute(sql)
Response.Redirect "?action=list"
end if
%>
<%
action=Request("action")
id=Request("id")
Username=Request("Username")
Password=Request("Password")
OLD_pass=Request("OLD_pass")
NEW_pass=Request("NEW_pass1")
if action="yes" Then
 set rs=server.createobject("adodb.recordset") 
if id="" then
   set rsCheck = conn.execute("select Username from Admin where Username='" & trim(Request.Form("Username")) & "'")
    if not (rsCheck.bof and rsCheck.eof) then
      response.write "<script language='javascript'>alert('" & trim(Request.Form("Username")) & "����Ա�����Ѵ��ڣ�');history.back(-1);</script>"
      response.end
     end if
	 set rsCheck=nothing
   sql="select * from Admin" 
   rs.open sql,conn,3,3
   rs.addnew 		'��Ӽ�¼�����ݱ�ĩ�� 
   rs("Username")=Username
   rs("Password")=md5(Password)
else
   sql="select * from Admin where id="&id&"" 
   rs.open sql,conn,1,2
   rs("Username")=Username
   rs("Password")=md5(NEW_pass)   
end if
 rs.update 		'�������ݱ��¼
 rs.close
set rs=nothing
 Response.Redirect "?action=list"
end if
%>
<html>
<head>
<title><%=sysConfig%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/main.css" rel="stylesheet" type="text/css">
<script LANGUAGE="Javascript">
function check()
{

  if (document.add.Username.value=="")
     {
      alert("�������ʻ�����")
      document.add.Username.focus()
      document.add.Username.select()
      return
     }
  if (document.add.Password1.value=="")
     {
      alert("�������������룡")
      document.add.Password1.focus()
      document.add.Password1.select()		
      return
     }
  if (document.add.Password1.value.length<6)
     {
      alert("����д����Ա����(�ַ�����6-16λ֮��)��")
      document.add.Password1.focus()
      document.add.Password1.select()		
      return
     }
  if (document.add.Password1.value.length>16)
     {
      alert("����д����Ա����(�ַ�����6-16λ֮��)��")
      document.add.Password1.focus()
      document.add.Password1.select()		
      return
     }
  if (document.add.Password1.value!=document.add.Password2.value)
     {
      alert("��������������벻��ͬ��")
      document.add.Password1.focus()
      document.add.Password1.select()		
      return
     }
     document.add.submit()
}
</script>
<script LANGUAGE="Javascript">
function check2()
{
  if (document.add.Username.value=="")
     {
      alert("�������ʻ�����")
      document.add.Username.focus()
      document.add.Username.select()
      return
     }
  if (document.add.NEW_pass1.value=="")
     {
      alert("�������������룡")
      document.add.NEW_pass1.focus()
      document.add.NEW_pass1.select()		
      return
     }
  if (document.add.NEW_pass1.value.length<6)
     {
      alert("����д����Ա����(�ַ�����6-16λ֮��)��")
      document.add.NEW_pass1.focus()
      document.add.NEW_pass1.select()		
      return
     }
  if (document.add.NEW_pass1.value.length>16)
     {
      alert("����д����Ա����(�ַ�����6-16λ֮��)��")
      document.add.NEW_pass1.focus()
      document.add.NEW_pass1.select()		
      return
     }
  if (document.add.NEW_pass1.value!=document.add.NEW_pass2.value)
     {
      alert("��������������벻��ͬ��")
      document.add.NEW_pass2.focus()
      document.add.NEW_pass2.select()		
      return
     }
     document.add.submit()
}
</script>

</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td bgcolor="#FFFFFF">
	<%if action="add" then%><br>
      <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
        <form name="add" method="post" action="Admin.asp">
        <tr align="center" bgcolor="#F2FDFF">
          <td colspan="2"  class="optiontitle">��ӹ���Ա</td>
        </tr>
        <tr align="center" bgcolor="#F2FDFF">
          <td width="10%" align="right">��½�ʺţ�</td>
          <td align="left"><input name="Username" type="text" id="Username"> <font color="red">*</font></td>
        </tr>
		<tr align='center' bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF"> ��½���룺</td>
		  <td align='left'><input name="Password1" type="password" id="Password1"> <font color="red">*</font>6-16���ַ�</td>
		</tr>
		<tr align='center' bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF"> ����ȷ�ϣ�</td>
		  <td align='left'><input name="Password2" type="password" id="Password2"> <font color="red">*</font></td>
		</tr>
        <tr align="center" bgcolor="#ebf0f7">
          <td  colspan="2" >
		     <INPUT TYPE="hidden" name="action" value="yes">
            <input type="button" name="Submit" value="�ύ" onClick="check()">
          	<input type="button" name="Submit2" value="����" onClick="history.back(-1)"></td>
        </tr>
		</FORM>
      </table> 
    <%end if%>
	
     <%if action="list" then%><br>
      <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
        <tr align="center" bgcolor="#F2FDFF">
          <td colspan="5"  class="optiontitle">����Ա�б�</td>
        </tr>
        <tr align="center">
          <td align="center"  bgcolor="#ebf0f7" width="5%">���</td>
          <td width="75%" colspan="2" align="center"  bgcolor="#ebf0f7">����Ա����</td>
          <td align="center"  bgcolor="#ebf0f7" width="20%">ִ�в���</td>
        </tr>
<%
set rs=server.createobject("adodb.recordset") 
sql="select * from Admin" 
rs.open sql,conn,1,1
do while not rs.eof
%>
        <tr align='center' bgcolor='#FFFFFF' onmouseover='this.style.background="#F2FDFF"' onmouseout='this.style.background="#FFFFFF"'>
		  <td width="5%"><%=rs("id")%></td>
          <td align='left'><%=rs("Username")%></td>
          <td align='left'><%=rs("Password")%></td>
          <td align='center'><%if rs("id")=1 then%><IMG src="images/edit.gif" align="absmiddle"><a href="?action=edit&id=<%=rs("id")%>">�޸�</a> | <IMG src="images/drop.gif" align="absmiddle">ɾ��</td><%else%><IMG src="images/edit.gif" align="absmiddle"><a href="?action=edit&id=<%=rs("id")%>">�޸�</a> | <IMG src="images/drop.gif" align="absmiddle"><a href="javascript:DoEmpty('?wor=del&id=<%=rs("id")%>')">ɾ��</a></td><%end if%>
        </tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
        <tr align="right" bgcolor="#ebf0f7">
          <td colspan="5"><A HREF="admin.asp?action=add">��ӹ���Ա</A></td>
        </tr>
      </table> 
<%
end if
if action="edit" then
set rs=server.createobject("adodb.recordset") 
sql="select * from Admin where id="&Request("id")
rs.open sql,conn,1,1
if not rs.eof Then
%>
<br>
	  <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
        <form name="add" method="post" action="admin.asp">
		<tr align="center" bgcolor="#F2FDFF">
		  <td colspan=2  class="optiontitle">�޸��û�</td>
		</tr>
		<tr align="center" bgcolor="#F2FDFF">
          <td width="10%" align="right">�û�����</td>
          <td align="left"><input name="Username" type="text" id="Username" value="<%=rs("Username")%>"></td>
        </tr>
		<tr align='center' bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF"> ��½���룺</td>
		  <td align='left'><input name="NEW_pass1" type="password" id="NEW_pass1"> <font color="red">*</font>6-16���ַ�</td>
		</tr>
		<tr align='center' bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF"> ����ȷ�ϣ�</td>
		  <td align='left'><input name="NEW_pass2" type="password" id="NEW_pass2"> <font color="red">*</font></td>
		</tr>
	    <tr align="center" bgcolor="#ebf0f7">
		  <td colspan="2">
		   <input type="hidden" name="action" value="yes"> <input type="button" name="Submit" value="�ύ" onClick="check2()">
           <input type="button" name="Submit2"value="����" onClick="history.back(-1)"> <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
		 </td>
	   </tr>
  	 </form>
  	</table>
<%
end if
end if
%>
    </td>
  </tr>
</table>
</body>
</html>