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
		sPictName = "‚©‚à‚ß‘å‹´"
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
		sPictName = "‘Ò‹@ê"
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
		sPictName = "ƒQ[ƒg‘O"
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
		sPictName = "ICCTƒQ[ƒg‘O"
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

' Logo—Í
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
Select Case sPICT
	Case "1"
		WriteLogM oFS, "Unknown", "8202", "Œg‘Ñ-‚©‚à‚ß‘å‹´‰f‘œ", "00",sPhoneType, ","
	Case "2"
		WriteLogM oFS, "Unknown", "8203", "Œg‘Ñ-‘Ò‹@ê‰f‘œ", "00",sPhoneType, ","
	Case "3"
		WriteLogM oFS, "Unknown", "8204", "Œg‘Ñ-ƒQ[ƒg‘O‰f‘œ", "00",sPhoneType, ","
	Case "4"
		WriteLogM oFS, "Unknown", "8205", "Œg‘Ñ-ICCTƒQ[ƒg‘O‰f‘œ", "00",sPhoneType, ","
End Select

    Dim fs, f, strUpdateTime, strPath, dateTimeTmp
	strPath = Server.MapPath(sPictPath)
    Set f = oFS.GetFile(strPath)
	dateTimeTmp = f.DateLastModified
	strUpdateTime = Year(dateTimeTmp) & "/" & _
		Right("0" & Month(dateTimeTmp), 2) & "/" & _
		Right("0" & Day(dateTimeTmp), 2) & " " & _
		Right("0" & Hour(dateTimeTmp), 2) & ":" & _
		Right("0" & Minute(dateTimeTmp), 2) & "<br>Œ»Ý"

Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb—pƒ^ƒO‚ð•ÒW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="<%=sPictName%>‰æ‘œ">
		<center>
		<%=sPictName%>‰f‘œ<br>
		<center>
		<img src="<%=sPictPath%>" alt="<%=sPictName%>"><br>
		<center>
		<%=strUpdateTime%><br>
		<center>
		<a task="gosub" dest="index.asp">ÒÆ­°</a>
	</display>
	</hdml>
<%
Else
	' EzWebˆÈŠO‚Ìƒ^ƒO‚ð•ÒW
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag(sPictName & "‰æ‘œ")%>
	</head>
	<body>
	<center>
	<%=sPictName%>‰f‘œ<br>
	<img src="<%=sPictPath%>" alt="<%=sPictName%>"><br>
	<%=strUpdateTime%>
	<form action="index.asp" method="get">
		<input type="submit" value="ÒÆ­°">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
