<!--#include file="inc/right.asp"--> 
<!--#include file="inc/conn.asp"-->
<%
keywords=request("keywords")
LX=request("LX")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>维度先生维修管理系统</title>
<link href="images/main.css" rel="stylesheet" type="text/css">
<script language=JavaScript>
<!--
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后此记录里的所有内容都将被删除并且无法恢复！"))
alert("删除信息成功\n 返回列表")
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
          <td class="optiontitle" colspan="10">查询报修信息</td>	
        </tr>
        <tr align="center" bgcolor="#F2FDFF">
          <form name="search" method="get" action="Search.asp">
            <td height="30" colspan="10"> <strong>查找类别：</strong><select name="LX">
          <option value="1">报修编号</option>1
          <option value="2">客户姓名</option>3
          <option value="3">联系电话</option>4
          <option value="4">联系QQ</option>5
          <option value="5">电脑品牌</option>2
		  <option value="6">电脑具体型号</option>6
		  <option value="7">故障描述</option>7
        </select> <input name="keywords" type="text" id="keywords" size="30"> 
              <input name="Query" type="submit" id="Query" value="查 询"></td>
          </form>
        </tr>
        <tr align="center" bgcolor="#ebf0f7">
          <td>报修编号</td>
          <td >客户姓名</td>
		  <td>联系电话</td>
		  <td>联系QQ</td>
		  <td >电脑品牌</td>
		  <td >电脑具体型号</td>
		  <td >故障描述</td>
		  <td >报修时间</td>
		  <td >报修状态</td>
          <td width="140">执行操作</td>
		</tr> 
<%
if keywords<>"" then
   sql="select * from Info where 5=5"
 if LX<>"" Then
   Select Case LX
     Case "1" 
	sql=sql & " and xx1 like '%"&keywords&"%' "
	 Case "2"
	sql=sql & " and xx3 like '%"&keywords&"%' "
     Case "3" 
	sql=sql & " and xx4 like '%"&keywords&"%' "
	 Case "4"
	sql=sql & " and xx5 like '%"&keywords&"%' "
     Case "5" 
	sql=sql & " and xx2 like '%"&keywords&"%' "
	 Case "6"
	sql=sql & " and xx6 like '%"&keywords&"%' "
	 Case "7"
	sql=sql & " and xx7 like '%"&keywords&"%' "
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
		  <td><%=rs("xx3")%></td>
          <td><%=rs("xx4")%></td>
		  <td><%=rs("xx5")%></td>
		  <td><%=rs("xx2")%></td>
		  <td><%=rs("xx6")%></td>
		  <td><%=rs("xx7")%></td>
		  <td><%=rs("addtime")%></td>
		  <td><%if rs("xx7")="待审核" then
		  c="red"
		  elseif rs("xx7")="维修中" then 
		  c="blue"
		  else c="green"
		  end if%><font color="<%=c%>"><%=rs("xx7")%></font></td>
          <td><IMG src="images/view.gif" align="absmiddle"><a href="Info.asp?action=view&id=<%=rs("id")%>">查看</a> | <IMG src="images/edit.gif" align="absmiddle"><a href="Info.asp?action=edit&id=<%=rs("id")%>">修改</a> | <IMG src="images/drop.gif" align="absmiddle"><a href="javascript:DoEmpty('Info.asp?wor=del&id=<%=rs("id")%>&action=list&ToPage=<%=intCurPage%>')">删除</a></td>
<%
rs.MoveNext 
next
%>

		</tr>
<%
else
%>
        <tr align="center" bgcolor="#FFFFFF">
          <td colspan="10">对不起！目前库中还没有 <font color="#FF0000"><%=keywords%></font> 信息！</td>
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