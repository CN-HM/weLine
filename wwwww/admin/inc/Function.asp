<%
'***************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��true ----�Ѿ���װ
'       false ----û�а�װ
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = false 
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = true 
	Set xTestObj = Nothing
	Err = 0
End Function
%>
