<!--#include file="inc/right.asp"--> 
<!--#include file="inc/conn.asp"-->
<%
keywords=request("keywords")
LX=request("LX")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ά������ά�޹���ϵͳ</title>
<link href="images/main.css" rel="stylesheet" type="text/css">
<script language=JavaScript>
<!--
function DoEmpty(params)
{
if (confirm("���Ҫɾ��������¼��ɾ����˼�¼����������ݶ�����ɾ�������޷��ָ���"))
alert("ɾ����Ϣ�ɹ�\n �����б�")
window.location = params ;
}
-->
</script>
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td bgcolor="#FFFFFF"><br>
      <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
		<tr align="center" bgcolor="#F2FDFF">
          <td class="optiontitle" colspan="10">��ѯ������Ϣ</td>	
        </tr>
        <tr align="center" bgcolor="#F2FDFF">
          <form name="search" method="get" action="SearchPerson.asp">
            <td height="30" colspan="10"> <strong>�������</strong><select name="LX">
          <option value="1">ά����Ա</option>1
          <option value="2">���ޱ��</option>3
        </select> <input name="keywords" type="text" id="keywords" size="30"> 
              <input name="Query" type="submit" id="Query" value="�� ѯ"></td>
          </form>
        </tr>
        <tr align="center" bgcolor="#ebf0f7">
          <td>���ޱ��</td>
		  <td >ά����Ա</td>
          <td >�ͻ�����</td>
		  <td >��������</td>
		  <td >����״̬</td>
		  <td >����ʱ��</td>
          <td width="140">ִ�в���</td>
		</tr> 
<%
if keywords<>"" then
   sql="select * from Info where 5=5"
 if LX<>"" Then
   Select Case LX
     Case "1" 
	sql=sql & " and adduser like '%"&keywords&"%' "
	 Case "2"
	sql=sql & " and xx1 like '%"&keywords&"%' "
   end select
   sql=sql & " order by id desc"
 end if
 set rs=server.createobject("adodb.recordset") 
 rs.open sql,conn,1,1
 if not rs.eof Then
    proCount=rs.recordcount
    For i = 1 to proCount
    if rs.EOF then     
    Exit For 
    end if
%>
		<tr align='center' bgcolor='#FFFFFF' onmouseover='this.style.background="#ebf0f7"' onmouseout='this.style.background="#FFFFFF"'>
          <td><%=rs("xx1")%></td>
		  <td><%=rs("adduser")%></td>
          <td><%=rs("xx3")%></td>
		  <td><%=rs("xx7")%></td>
		  <td><%if rs("xx8")="�����" then
		  c="red"
		  elseif rs("xx8")="ά����" then 
		  c="blue"
		  else c="green"
		  end if%><font color="<%=c%>"><%=rs("xx8")%></font></td>
		  <td><%=rs("addtime")%></td>
		  
          <td><IMG src="images/view.gif" align="absmiddle"><a href="Info.asp?action=view&id=<%=rs("id")%>">�鿴</a> | <IMG src="images/edit.gif" align="absmiddle"><a href="Info.asp?action=edit&id=<%=rs("id")%>">�޸�</a> | <IMG src="images/drop.gif" align="absmiddle"><a href="javascript:DoEmpty('Info.asp?wor=del&id=<%=rs("id")%>&action=list&ToPage=<%=intCurPage%>')">ɾ��</a></td>
<%
rs.MoveNext 
next
%>

		</tr>
<%
else
%>
        <tr align="center" bgcolor="#FFFFFF">
          <td colspan="10">�Բ���Ŀǰ���л�û�� <font color="#FF0000"><%=keywords%></font> ��Ϣ��</td>
        </tr>
<%
end if
rs.close
set rs=nothing
end if
%>
      </table> 
    </td>
  </tr>
</table>
</body>
</html>