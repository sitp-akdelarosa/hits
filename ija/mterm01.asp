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
	sErrMsg = "DBエラー"
Else
	sRecWait = ""
	sDelWait = ""
	sRDWait = ""
	If Not IsNull(rs("RecWaitTime")) Then
		If rs("RecWaitTime")>120 Then
			sRecWait = "---"
		Else
			sRecWait = CStr(rs("RecWaitTime")) & "分"
		End If
	End If
	If Not IsNull(rs("DelWaitTime")) Then
		If rs("DelWaitTime")>120 Then
			sDelWait = "---"
		Else
			sDelWait = CStr(rs("DelWaitTime")) & "分"
		End If
	End If
	If Not IsNull(rs("RDWaitTime")) Then
		If rs("RDWaitTime")>120 Then
			sRDWait = "---"
		Else
			sRDWait = CStr(rs("RDWaitTime")) & "分"
		End If
	End If
End If
rs.Close

'ADD START HiTS Ver2 By SEIKO N.Ooshige
dim IcInTime,IcOutTime
sSQL = "SELECT RecWaitTime, DelInWaitTime,DelOutWaitTime FROM Terminal2 WHERE Terminal='IC'"
rs.Open sSQL, conn, 0, 1, 1
If rs.Eof Then
	sErrMsg = "DBエラー"
Else
	IcInTime=""
	IcOutTime=""
	If Not IsNull(rs("RecWaitTime")) Then
		If rs("RecWaitTime")<2 or rs("RecWaitTime")>240 Then
			IcInTime = "---"
		Else
			IcInTime = CStr(rs("RecWaitTime")) & "分"
		End If
	End If
	If Not IsNull(rs("DelInWaitTime")) AND Not IsNull(rs("DelOutWaitTime")) Then
		IcOutTime = rs("DelInWaitTime") + rs("DelOutWaitTime")
		If IcOutTime<2 or IcOutTime>240 Then
			IcOutTime = "---"
		Else
			IcOutTime = IcOutTime & "分"
		End If
	End If
End If
rs.Close
'ADD END HiTS Ver2 By SEIKO N.Ooshige
conn.Close

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "8201", "携帯-ゲート内所要時間", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="ターミナル情報">
		<center>
		【ﾀ-ﾐﾅﾙ情報】<br>
		<center>
		香　椎<br>
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
			ﾀ-ﾐﾅﾙ内所要時間<br>
			<center>
			搬入のみ…<%=sRecWait%><br>
			<center>
			搬出のみ…<%=sDelWait%><br>
			<center>
			搬出入…<%=sRDWait%><br>
			<br><br>
			<center>
			アイランドシティ<br>
			<center>
			ﾀ-ﾐﾅﾙ内所要時間<br>
			<center>
			搬入…<%=IcInTime%><br>
			<center>
			搬出…<%=IcOutTime%><br>
<%
		End If
%>
		<center>
		<a task="gosub" dest="index.asp">ﾒﾆｭｰ</a>
	</display>
	</hdml>
<%
Else
	' EzWeb以外のタグを編集
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("ターミナル情報")%>
	</head>
	<body>
	<center>
	【ﾀ-ﾐﾅﾙ情報】<br>
	香　椎<br>
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
				ﾀ-ﾐﾅﾙ内所要時間<br>
				搬入のみ…<%=sRecWait%><br>
				搬出のみ…<%=sDelWait%><br>
				搬出入…<%=sRDWait%><br>
			</td></tr>
		</table>
	<br><br>
	<center>
	アイランドシティ<br>
	<hr>
		<table border="0">
			<tr><td>
				ﾀ-ﾐﾅﾙ内所要時間<br>
				搬入…<%=IcInTime%><br>
				搬出…<%=IcOutTime%><br>
			</td></tr>
		</table>
<%
	End If
%>
	<form action="index.asp" method="get">
		<input type="submit" value="ﾒﾆｭｰ">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
