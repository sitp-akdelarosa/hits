<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common_cam.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()
' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "2403", "�g��-BL�ԍ��Ɖ�i�����ӓ��j", "00" , sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry key="blno" title="BL�ԍ��Ɖ�">
		<action type="accept" task="go" dest="mblno02cam.asp?blno=$blno">
		<center>
		�yBL�ԍ��Ɖ�z<br><br>
		BL�ԍ��F
	</entry>

	</hdml>
<%
Else
	' EzWeb�ȊO�̃^�O��ҏW
%>
	<html>
	
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("BL�ԍ��Ɖ�")%>
	</head>
	
	<body>
	<center>
	�yBL�ԍ��Ɖ�z
	<hr>
	<form action="mblno02cam.asp" METHOD="get">
		BL�ԍ�����<br>
		<input type="text" name="BLno" maxlength="20" <%=GetTextSizeMode(20, "A")%>><br>
		<br>
		<input type="submit" value="����" >
	</form>
	<hr>

	<!--<br><a href="../mblno01.asp">���ŁEICCT�Ɖ�</a>-->

	</body>
	
	</html>
<%
End If
%>
