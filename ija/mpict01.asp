<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim vPICT, sPICT
Dim sPictPath, sPictName
Dim sPhoneType
sPhoneType = GetPhoneType()

vPICT = Trim(Request.QueryString("PICT"))
If IsEmpty(vPICT) Then
	sPICT = "1"
Else
	sPICT = Trim(vPICT)
End If

Select Case sPICT
	Case "1"
		sPictName = "�����ߑ勴"
		Select Case sPhoneType
			Case "I"
				sPictPath = "i-kamome.gif"
			Case "J"
				sPictPath = "j-kamome.png"
			Case "E"
				sPictPath = "e-kamome.png"
			Case Else
				sPictPath = "e-kamome.png"
		End Select
	Case "2"
		sPictName = "�ҋ@��"
		Select Case sPhoneType
			Case "I"
				sPictPath = "i-taiki.gif"
			Case "J"
				sPictPath = "j-taiki.png"
			Case "E"
				sPictPath = "e-taiki.png"
			Case Else
				sPictPath = "e-taiki.png"
		End Select
	Case "3"
		sPictName = "�Q�[�g�O"
		Select Case sPhoneType
			Case "I"
				sPictPath = "i-gate.gif"
			Case "J"
				sPictPath = "j-gate.png"
			Case "E"
				sPictPath = "e-gate.png"
			Case Else
				sPictPath = "e-gate.png"
		End Select
	Case "4"
		sPictName = "ICCT�Q�[�g�O"
		Select Case sPhoneType
			Case "I"
				sPictPath = "i-gate.icct.gif"
			Case "J"
				sPictPath = "j-gate.icct.png"
			Case "E"
				sPictPath = "e-gate.icct.png"
			Case Else
				sPictPath = "e-gate.icct.png"
		End Select
	Case Else
		sPictName = ""
		sPictPath = ""
End Select

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
Select Case sPICT
	Case "1"
		WriteLogM oFS, "Unknown", "8202", "�g��-�����ߑ勴�f��", "00",sPhoneType, ","
	Case "2"
		WriteLogM oFS, "Unknown", "8203", "�g��-�ҋ@��f��", "00",sPhoneType, ","
	Case "3"
		WriteLogM oFS, "Unknown", "8204", "�g��-�Q�[�g�O�f��", "00",sPhoneType, ","
	Case "4"
		WriteLogM oFS, "Unknown", "8205", "�g��-ICCT�Q�[�g�O�f��", "00",sPhoneType, ","
End Select

    Dim fs, f, strUpdateTime, strPath, dateTimeTmp
	strPath = Server.MapPath(sPictPath)
    Set f = oFS.GetFile(strPath)
	dateTimeTmp = f.DateLastModified
	strUpdateTime = Year(dateTimeTmp) & "/" & _
		Right("0" & Month(dateTimeTmp), 2) & "/" & _
		Right("0" & Day(dateTimeTmp), 2) & " " & _
		Right("0" & Hour(dateTimeTmp), 2) & ":" & _
		Right("0" & Minute(dateTimeTmp), 2) & "<br>����"

Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="<%=sPictName%>�摜">
		<center>
		<%=sPictName%>�f��<br>
		<center>
		<img src="<%=sPictPath%>" alt="<%=sPictName%>"><br>
		<center>
		<%=strUpdateTime%><br>
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
		<%=GetTitleTag(sPictName & "�摜")%>
	</head>
	<body>
	<center>
	<%=sPictName%>�f��<br>
	<img src="<%=sPictPath%>" alt="<%=sPictName%>"><br>
	<%=strUpdateTime%>
	<form action="index.asp" method="get">
		<input type="submit" value="�ƭ�">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
