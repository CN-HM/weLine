<!--#include file="inc/right.asp"--> 
<!--#include file="inc/conn.asp"-->
<%
if Request("wor")="del" then
id=request("id")
idArr=split(id,",")
for i=0 to ubound(idArr)
sql="delete from Info where id="&trim(idArr(i))
conn.execute(sql)
next
end if
%>
<%
action=Request("action")
id=Request("id")
if action="yes" Then
   set rs=server.createobject("adodb.recordset") 
	if id="" then
	   set rsCheck = conn.execute("select xx1 from Info where xx1='" & trim(Request.Form("xx1")) & "'")
	   'rsCheck.bof and rsCheck.eof  如果数据集的头跟尾相同则为1   前面NOT则取反为0  则不执行下面报错语句
		 if not (rsCheck.bof and rsCheck.eof) then
		  response.write "<script language='javascript'>alert('报修编号 " & trim(Request.Form("xx1")) & " 已存在，请检查！');history.back(-1);</script>"
		  response.end
		 end if
	   set rsCheck=nothing
	   sql="select * from Info" 
	   rs.open sql,conn,3,3
	   rs.addnew
	else
	   sql="select * from Info where id="&id&"" 
	   rs.open sql,conn,1,2
	   
	end if
for i=1 to 8
rs("xx"&i)=Request("xx"&i)
next
rs("addtime")=Request("addtime")
rs("adduser")=session("admin_name")
 rs.update
 rs.close
set rs=nothing
 Response.Redirect "?action=list"
 
end if
%>
<html>
<head>
<title>维度先生维修管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/main.css" rel="stylesheet" type="text/css">
<script language=JavaScript type=text/JavaScript>

function CheckAll() {
 var checkboxs=document.getElementsByName("id");
 for (var i=0;i<checkboxs.length;i++) {
  var e=checkboxs[i];
  e.checked=!e.checked;
 }
}

</script>
<script language=JavaScript>
<!--
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后此记录里的所有内容都将被删除并且无法恢复！"))
window.location = params ;
}

function check()
{

  if (document.add.xx3.value=="")
     {
      alert("请填写客户姓名！")
      document.add.xx3.focus()
      document.add.xx3.select()
      return
     }
	 
  if (document.add.xx4.value=="")
     {
      alert("请填写联系电话！")
      document.add.xx4.focus()
      document.add.xx4.select()
      return
     }
  
  if (document.add.xx5.value=="")
     {
      alert("请填写联系QQ！")
      document.add.xx5.focus()
      document.add.xx5.select()
      return
     }
	 
  if (document.add.xx6.value=="")
     {
      alert("请填写电脑具体型号！")
      document.add.xx6.focus()
      document.add.xx6.select()
      return
     }
	 
  if (document.add.xx6.value=="")
     {
      alert("请填写维修问题描述！")
      document.add.xx7.focus()
      document.add.xx7.select()
      return
     }
	 	 
     document.add.submit()
}

 function next()
 {
  if(event.keyCode==13)event.keyCode=9;
 }
-->
</script>
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td bgcolor="#FFFFFF">
	<%if action="add" then%><BR>
        <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
        <form name="add" method="post" action="Info.asp">
        <tr align="center" bgcolor="#F2FDFF">
          <td colspan="2"  class="optiontitle"> 添加报修信息 </td>
        </tr>	
		<tr bgcolor="#FFFFFF">
	   <%
	   Randomize
	   bh=year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())+int(999*rnd())%>
		  <td width="13%" align="right">报修编号：</td>
		  <td ><input name="xx1" type="text" id="xx1" onKeyDown="next()" value="W<%=bh%>" readonly /> </td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>客户姓名：</td>
		  <td  ><input name="xx3" type="text" id="xx3" onKeyDown="next()"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>联系电话：</td>
		  <td  ><input name="xx4" type="text" id="xx4" onKeyDown="next()"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>联系QQ：</td>
		  <td  ><input name="xx5" type="text" id="xx5" onKeyDown="next()"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'> 电脑品牌：</td>
          <td><%
Set rsSclass = Server.CreateObject("ADODB.Recordset")
rsSclass.open "select * from Sclass",Conn,1,2
%>
		  <SELECT name='xx2' id="xx2">
		  <%
		  do while not rsSclass.eof
		  %>
		  <option value="<%=rsSclass("title")%>"><%=rsSclass("title")%></option>
          <%
			rsSclass.movenext
			loop
			set rsSclass=nothing
		  %>
		  </SELECT></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>电脑具体型号：</td>
		  <td  ><input name="xx6" type="text" id="xx6" onKeyDown="next()"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF"> 维修问题描述：</td>
		  <td ><textarea name="xx7" cols="60" rows="5" id="xx7" onKeyDown="next()"></textarea></td>
		</tr>
				
	    <tr bgcolor='#FFFFFF'>
		  <td align='right' >报修时间：</td>
		  <td ><input name="addtime" type="text" id="addtime" value="<%=now()%>" onKeyDown="next()" /></td>
		 </tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'> 报修状态：</td>
          <td>
           <select name="xx8" id="xx8">
		   <option value="待审核">待审核</option>
		   <option value="维修中">维修中</option>
		   <option value="已维修">已维修</option>
            </select></td>
		</tr>
		
        <tr align="center" bgcolor="#ebf0f7">
          <td colspan="2" >
		    <input TYPE="hidden" name="action" value="yes"/>
            <input type="button" name="Submit" value="提交" onClick="check()"/>
          	<input type="button" name="Submit2" value="返回" onClick="history.back(-1)"/></td>
        </tr>
		</FORM>
      </table> 
<%end if%>
<br>
<%if action="list" then%>
      <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
        <tr align="center" bgcolor="#F2FDFF">
          <td colspan="12"  class="optiontitle">报修信息</td>
        </tr>
<%
TN=trim(request("TN"))
%>
          <form name="search" method="post" action="info.asp?action=list">
        <tr align="center" bgcolor="#F2FDFF">
          <td colspan="12">
            <input name="TN" type="radio" value="" <%if TN="" then response.Write("checked") end if%> onClick="javascript:this.form.submit();"/>所有信息
	        <input type="radio" name="TN" value="待审核" <%if TN="待审核" then response.Write("checked") end if%> onClick="javascript:this.form.submit();"/>待审核
	        <input type="radio" name="TN" value="维修中" <%if TN="维修中" then response.Write("checked") end if%> onClick="javascript:this.form.submit();"/>维修中
	        <input type="radio" name="TN" value="已维修" <%if TN="已维修" then response.Write("checked") end if%> onClick="javascript:this.form.submit();"/>已维修
		  </td>
        </tr>
        <tr align="center" bgcolor="#ebf0f7">
		  <td width="35">选中</td>
          <td>报修编号</td>
          <td >客户姓名</td>
		  <td>联系电话</td>
		  <td>联系QQ</td>
		  <td >电脑品牌</td>
		  <td >电脑具体型号</td>
		  <td >故障描述</td>
		  <td >报修时间</td>
		  <td >维修人员</td>
		  <td >报修状态</td>
          <td width="140">执行操作</td>
        </tr>	
<%
 sql="select * from Info where 5=5"
 if TN<>"" then
	sql=sql & " and xx8='"&TN&"' "
 end if
 sql=sql & " order by id desc"
 set rs=server.createobject("adodb.recordset") 
 rs.open sql,conn,1,1
 if not rs.eof then
 proCount=rs.recordcount
	rs.PageSize=50
     if not IsEmpty(Request("ToPage")) then
	    ToPage=CInt(Request("ToPage"))
		if ToPage>rs.PageCount then
		   rs.AbsolutePage=rs.PageCount
		   intCurPage=rs.PageCount
		elseif ToPage<=0 then
		   rs.AbsolutePage=1
		   intCurPage=1
		else
		   rs.AbsolutePage=ToPage
		   intCurPage=ToPage
		end if
	 else
		rs.AbsolutePage=1
		intCurPage=1
	 end if
	 intCurPage=CInt(intCurPage)
	 For i = 1 to rs.PageSize
	 if rs.eof then     
	 Exit For 
	 end if
%>
       <form name="del" action="" method="post">
        <tr align='center' bgcolor='#FFFFFF' onmouseover='this.style.background="#F2FDFF"' onmouseout='this.style.background="#FFFFFF"'>
          <td><input type="checkbox" name="id" value="<%=rs("id")%>"/></td>
          <td><%=rs("xx1")%></td>
		  <td><%=rs("xx3")%></td>
          <td><%=rs("xx4")%></td>
		  <td><%=rs("xx5")%></td>
		  <td><%=rs("xx2")%></td>
		  <td><%=rs("xx6")%></td>
		  <td><%=rs("xx7")%></td>
		  <td><%=rs("addtime")%></td>
		  <td><%=rs("adduser")%></td>
		  <td>
		  <%if rs("xx8")="待审核" then
		  c="red"
		  elseif rs("xx8")="维修中" then 
		  c="blue"
		  else c="green"
		  end if%><font color="<%=c%>"><%=rs("xx8")%></font></td>
          <td><IMG src="images/view.gif" align="absmiddle"><a href="?action=view&id=<%=rs("id")%>">查看</a> | <IMG src="images/edit.gif" align="absmiddle"><a href="?action=edit&id=<%=rs("id")%>">修改</a> | <IMG src="images/drop.gif" align="absmiddle"><a href="javascript:DoEmpty('?wor=del&id=<%=rs("id")%>&action=list&ToPage=<%=intCurPage%>')">删除</a></td>
        </tr>
<%
rs.movenext 
next
%>
		<tr bgcolor="#ffffff">
		  <td colspan="12">&nbsp;&nbsp;
		   <input name="chkall" type="checkbox" id="chkall" value="select" onclick="CheckAll()" /> <a>全选</a>
		   <input name="wor" type="hidden" id="wor" value="del" />
		   <input type="submit" name="Submit3" value="删除所选" onClick="{if(confirm('确定要删除记录吗？删除后将被无法恢复！')){return true;}return false;}" />
		  </td>
		</tr>
		
		</form>
        <tr align="center" bgcolor="#ebf0f7">
          <td colspan="12"> 总共：<font color="#ff0000"><%=rs.PageCount%></font>页, <font color="#ff0000"><%=proCount%></font>条报修信息, 当前页：<font color="#ff0000"><%=intCurPage%> </font><%if intCurPage<>1 then%><a href="?action=list&TN=<%=TN%>">首页</a> | <a href="?action=list&TN=<%=TN%>&ToPage=<%=intCurPage-1%>">上一页</a> | <% end if
if intCurPage<>rs.PageCount then %><a href="?action=list&TN=<%=TN%>&ToPage=<%=intCurPage+1%>">下一页</a> | <a href="?action=list&TN=<%=TN%>&ToPage=<%=rs.PageCount%>"> 最后页</a><% end if%></span></td>
        </tr>
<%
else
%>
        <tr align="center" bgcolor="#ffffff">
          <td colspan="12">对不起！目前数据库中还没有添加报修信息！</td>
        </tr>
<%
rs.close
set rs=nothing
end if
%>
      </table>
<br>
<%end if%>
<%if action="edit" then
set rs=server.createobject("adodb.recordset") 
sql="select * from Info where id="&Request("id")
rs.open sql,conn,1,1
if not rs.eof Then
%>
	  <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
       <form name="add" method="post" action="Info.asp">
		<tr align="center" bgcolor="#F2FDFF">
		  <td colspan=2  class="optiontitle"> 修改报修信息 </td>
		</tr>
		<tr bgcolor='#F2FDFF'>
		  <td width="10%" align="right">报修编号：</td>
		  <td ><input name="xx1" type="text" id="xx1" onKeyDown="next()" value="<%=rs("xx1")%>" readonly /> </td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>客户姓名：</td>
		  <td  ><input name="xx3" type="text" id="xx3" onKeyDown="next()" value="<%=rs("xx3")%>"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>联系电话：</td>
		  <td  ><input name="xx4" type="text" id="xx4" onKeyDown="next()" value="<%=rs("xx4")%>"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>联系QQ：</td>
		  <td  ><input name="xx5" type="text" id="xx5" onKeyDown="next()" value="<%=rs("xx5")%>"/></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'> 电脑品牌：</td>
          <td><%
Set rsSclass = Server.CreateObject("ADODB.Recordset")
rsSclass.open "select * from Sclass",Conn,1,2
%>
		  <SELECT name='xx2' id="xx2">
		  <%
		  do while not rsSclass.eof
		  %>
		  <option value="<%=rsSclass("title")%>" <%if rs("xx2")=rsSclass("title") then response.write("Selected")%>><%=rsSclass("title")%></option>
          <%
			rsSclass.movenext
			loop
			set rsSclass=nothing
		  %>
		  </SELECT></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>电脑具体型号：</td>
		  <td  ><input name="xx6" type="text" id="xx6" onKeyDown="next()" value="<%=rs("xx6")%>" /></td>
		</tr>
		
		<tr bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF"> 维修问题描述：</td>
		  <td ><textarea name="xx7" cols="60" rows="5" id="xx7" onKeyDown="next()"><%=rs("xx7")%></textarea></td>
		</tr>
				
	    <tr bgcolor='#FFFFFF'>
		  <td align='right' >报修时间：</td>
		  <td ><input name="addtime" type="text" id="addtime" value="<%=rs("addtime")%>" onKeyDown="next()" /></td>
		 </tr>

		 <tr bgcolor='#FFFFFF'>
		  <td align='right'> 报修状态：</td>
          <td>
           <select name="xx8" id="xx8">
		   <option value="待审核" <%if rs("xx8")="待审核" then response.write("Selected")%>>待审核</option>
		   <option value="维修中" <%if rs("xx8")="维修中" then response.write("Selected")%>>维修中</option>
		   <option value="已维修" <%if rs("xx8")="已维修" then response.write("Selected")%>>已维修</option>
            </select></td>
		</tr>
		
		<tr align="center" bgcolor="#ebf0f7">
		  <td colspan="2">
		  <input type="hidden" name="action" value="yes"/>
          <input type="button" name="Submit2" value="提交" onClick="check()"/>
          <input type="button" name="Submit2" value="返回" onClick="history.back(-1)"/>
		  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"/></td>
		</tr>
  		</FORM>
  	</table>
<%
end if
end if
%>  
<%if action="view" then
set rs=server.createobject("adodb.recordset") 
sql="select * from Info where id="&Request("id")
rs.open sql,conn,1,1
if not rs.eof Then
%>
	  <table width="96%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#aec3de">
		<tr align="center" bgcolor="#F2FDFF">
		  <td colspan=2  class="optiontitle"> 查看报修信息 </td>
		</tr>
		<tr bgcolor='#F2FDFF'>
		  <td width="10%" align="right">报修编号：</td>
		  <td><%=rs("xx1")%></td>
		</tr>		
		<tr bgcolor='#FFFFFF'>
		  <td align='right'> 客户姓名：</td>
          <td><%=rs("xx3")%></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'> 联系电话：</td>
		  <td><%=rs("xx4")%></td>
		</tr>
        <tr bgcolor="#FFFFFF">
          <td align='right'> 联系QQ：</td>
          <td><%=rs("xx5")%></td>
        </tr>	
		<tr bgcolor='#FFFFFF'>
		  <td align='right' >电脑品牌：</td>
		  <td><%=rs("xx2")%></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right' >电脑具体型号：</td>
		  <td><%=rs("xx6")%></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right' >维修问题描述：</td>
		  <td><%=rs("xx7")%></td>
		</tr>		
	    <tr bgcolor='#FFFFFF'>
		  <td align='right' bgcolor="#FFFFFF">报修时间：</td>
		  <td><%=rs("addtime")%></td>
		</tr>
		<tr align="center" bgcolor="#ebf0f7">
		  <td colspan="2">
          <input type="button" name="Submit2" value="返回" onClick="history.back(-1)"/></td>
		</tr>
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