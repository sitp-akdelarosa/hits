<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "2201", "�g��-�R���e�i�ԍ��Ɖ�", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry name="p1" key="cont_e" format="*A" title="�R���e�i�ԍ��Ɖ�">
		<action type="accept" task="go" dest="#p2">
		<center>
		�y���Ŕԍ��Ɖ�z<br><br>
		�擪�p��4��:
	</entry>
	
	<entry name="p2" key="cont_s" format="*N">
		<action type="accept" task="go" dest="mcont02.asp?cont_e=$cont_e&cont_s=$cont_s">
		<center>
		�y���Ŕԍ��Ɖ�z<br><br>
		��������7��:
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
		<%=GetTitleTag("�R���e�i�ԍ��Ɖ�")%>
	</head>
	<body>
	<center>
	�y���Ŕԍ��Ɖ�z
	<hr>
	<form action="mcont02.asp" method="get">
		���Ŕԍ�����<br>
		<table border="0">
			<tr><td>
				�p��4��:
				<input type="text" name="cont_e" maxlength="4" <%=GetTextSizeMode(4, "A")%>><br>
			</td></tr>
			<tr><td>
				����:
				<input type="text" name="cont_s" maxlength="8" <%=GetTextSizeMode(8, "N")%>><br>
			</td></tr>
		</table>
		<input type="submit" value="����">
	</form>
	<hr>

	<br><a href="./cam/mcont01cam.asp">�����ӓ��Ɖ��</a>

	</body>
	</html>
<%
End If
%>
