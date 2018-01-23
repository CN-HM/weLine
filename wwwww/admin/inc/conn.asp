<%
'session.timeout=6000
'Server.ScriptTimeOut=2000

dim conn,rs
set conn=server.createobject("adodb.connection")
conn.open "PROVIDER=SQLOLEDB;DATA SOURCE=WIN-TRU9BJVUPPA;UID=sa;PWD=qazpl2010.;DATABASE=weidu"
Set rs = Server.CreateObject("ADODB.Recordset")
%>
