<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Const TERMINAL_CODE = "KA"

Dim sRecWait, sDelWait, sRDWait
Dim sSQL
Dim conn, rs
Dim sErrMsg

sErrMsg = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

ConnectSvr conn, rs

sSQL = "SELECT Terminal.RecWaitTime, Terminal.DelWaitTime, Terminal.RDWaitTime " & _
	" FROM Terminal WHERE Terminal.Terminal = '" & TERMINAL_CODE & "'"
rs.Open sSQL, conn, 0, 1, 1
If rs.Eof Then
	sErrMsg = "DB�G���["
Else
	sRecWait = ""
	sDelWait = ""
	sRDWait = ""
	If Not IsNull(rs("RecWaitTime")) Then
		If rs("RecWaitTime")>120 Then
			sRecWait = "---"
		Else
			sRecWait = CStr(rs("RecWaitTime")) & "��"
		End If
	End If
	If Not IsNull(rs("DelWaitTime")) Then
		If rs("DelWaitTime")>120 Then
			sDelWait = "---"
		Else
			sDelWait = CStr(rs("DelWaitTime")) & "��"
		End If
	End If
	If Not IsNull(rs("RDWaitTime")) Then
		If rs("RDWaitTime")>120 Then
			sRDWait = "---"
		Else
			sRDWait = CStr(rs("RDWaitTime")) & "��"
		End If
	End If
End If
rs.Close

'ADD START HiTS Ver2 By SEIKO N.Ooshige
dim IcInTime,IcOutTime
sSQL = "SELECT RecWaitTime, DelInWaitTime,DelOutWaitTime FROM Terminal2 WHERE Terminal='IC'"
rs.Open sSQL, conn, 0, 1, 1
If rs.Eof Then
	sErrMsg = "DB�G���["
Else
	IcInTime=""
	IcOutTime=""
	If Not IsNull(rs("RecWaitTime")) Then
		If rs("RecWaitTime")<2 or rs("RecWaitTime")>240 Then
			IcInTime = "---"
		Else
			IcInTime = CStr(rs("RecWaitTime")) & "��"
		End If
	End If
	If Not IsNull(rs("DelInWaitTime")) AND Not IsNull(rs("DelOutWaitTime")) Then
		IcOutTime = rs("DelInWaitTime") + rs("DelOutWaitTime")
		If IcOutTime<2 or IcOutTime>240 Then
			IcOutTime = "---"
		Else
			IcOutTime = IcOutTime & "��"
		End If
	End If
End If
rs.Close
'ADD END HiTS Ver2 By SEIKO N.Ooshige
conn.Close

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "8201", "�g��-�Q�[�g�����v����", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="�^�[�~�i�����">
		<center>
		�y�-��ُ��z<br>
		<center>
		���@��<br>
<%
		If sErrMsg <> "" Then
%>
			<br>
			<center>
			<%=sErrMsg%><br><br>
<%
		Else
%>
			<center>
			�-��ٓ����v����<br>
			<center>
			�����̂݁c<%=sRecWait%><br>
			<center>
			���o�̂݁c<%=sDelWait%><br>
			<center>
			���o���c<%=sRDWait%><br>
			<br><br>
			<center>
			�A�C�����h�V�e�B<br>
			<center>
			�-��ٓ����v����<br>
			<center>
			�����c<%=IcInTime%><br>
			<center>
			���o�c<%=IcOutTime%><br>
<%
		End If
%>
		<center>
		<a task="gosub" dest="index.asp">�ƭ�</a>
	</display>
	</hdml>
<%
Else
	' EzWeb�ȊO�̃^�O��ҏW
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("�^�[�~�i�����")%>
	</head>
	<body>
	<center>
	�y�-��ُ��z<br>
	���@��<br>
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<br>
		<%=sErrMsg%><br><br>
<%
	Else
%>
		<table border="0">
			<tr><td>
				�-��ٓ����v����<br>
				�����̂݁c<%=sRecWait%><br>
				���o�̂݁c<%=sDelWait%><br>
				���o���c<%=sRDWait%><br>
			</td></tr>
		</table>
	<br><br>
	<center>
	�A�C�����h�V�e�B<br>
	<hr>
		<table border="0">
			<tr><td>
				�-��ٓ����v����<br>
				�����c<%=IcInTime%><br>
				���o�c<%=IcOutTime%><br>
			</td></tr>
		</table>
<%
	End If
%>
	<form action="index.asp" method="get">
		<input type="submit" value="�ƭ�">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
