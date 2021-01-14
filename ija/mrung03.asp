<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
' �g�����U�N�V�����t�@�C�����o�͂���t�H���_
Const SEND_FOLDER = "send"

' �g�����U�N�V�����t�@�C���̊g���q
Const SEND_EXTENT = "snd"

' �����ꏊ
Const TRANS_PLACE = ""

Dim sUserID
Dim sContNo
Dim sOperation
Dim sOpeName
Dim tNow
Dim sDailyTransNo
Dim sRecord()
Dim sDateTime1, sDateTime2
Dim conn, rs
Dim sSQL
Dim sVoyage, sShipLine
Dim sErrMsg
Dim i,j

i = 0
j = 0

sErrMsg = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

sUserID = Trim(Request.QueryString("UserID"))
sContNo = Trim(Request.QueryString("ContNo"))
sOperation = Trim(Request.QueryString("operation"))

sDailyTransNo = GetDailyTransNo()
tNow = Now
sDateTime1 = ArrangeNum(Year(tNow), 4) & ArrangeNum(Month(tNow), 2) & ArrangeNum(Day(tNow), 2) & _
			ArrangeNum(Hour(tNow), 2) & ArrangeNum(Minute(tNow), 2)
sDateTime2 = sDateTime1 & ArrangeNum(Second(tNow), 2)

ConnectSvr conn, rs

ReDim sRecord(0)
sRecord(0) = ""
Select Case sOperation
	Case "A":
		sOpeName = "�����q�ɒ�"
		sSQL = "SELECT ImportCont.VslCode, ImportCont.ContNo, ImportCont.BLNo, VslSchedule.DsVoyage " & _
			" FROM ImportCont, VslSchedule " & _
			" WHERE ImportCont.ContNo = '" & sContNo & "' " & _
			" AND VslSchedule.VslCode=*ImportCont.VslCode " & _
			" AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl "
		rs.Open sSQL, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "�Ώۺ��łȂ�"
		Else
			Do While Not rs.Eof
				ReDim Preserve sRecord(i)

				If IsNull(rs("DsVoyage")) Then
					sVoyage = ""
				Else
					sVoyage = Trim(rs("DsVoyage"))
				End If
				sRecord(i) = sRecord(i) & sDailyTransNo & ","
				sRecord(i) = sRecord(i) & "IM10,"
				sRecord(i) = sRecord(i) & "R,"
				sRecord(i) = sRecord(i) & sDateTime2 & ","
				sRecord(i) = sRecord(i) & sUserID & ","
				sRecord(i) = sRecord(i) & TRANS_PLACE & ","
				sRecord(i) = sRecord(i) & Trim(rs("VslCode")) & ","
				sRecord(i) = sRecord(i) & sVoyage & ","
				sRecord(i) = sRecord(i) & Trim(rs("ContNo")) & ","
				sRecord(i) = sRecord(i) & Trim(rs("BLNo")) & ","
				sRecord(i) = sRecord(i) & sDateTime1

				rs.MoveNext
				i = i+1
			Loop
		End If
		rs.Close

	Case "B":
		sOpeName = "�f�o����"
		sSQL = "SELECT ImportCont.VslCode, ImportCont.ContNo, ImportCont.BLNo, VslSchedule.DsVoyage " & _
			" FROM ImportCont, VslSchedule " & _
			" WHERE ImportCont.ContNo = '" & sContNo & "' " & _
			" AND VslSchedule.VslCode=*ImportCont.VslCode " & _
			" AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl "
		rs.Open sSQL, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "�Ώۺ��łȂ�"
		Else
			Do While Not rs.Eof
				ReDim Preserve sRecord(i)

				If IsNull(rs("DsVoyage")) Then
					sVoyage = ""
				Else
					sVoyage = Trim(rs("DsVoyage"))
				End If
				sRecord(i) = sRecord(i) & sDailyTransNo & ","
				sRecord(i) = sRecord(i) & "IM11,"
				sRecord(i) = sRecord(i) & "R,"
				sRecord(i) = sRecord(i) & sDateTime2 & ","
				sRecord(i) = sRecord(i) & sUserID & ","
				sRecord(i) = sRecord(i) & TRANS_PLACE & ","
				sRecord(i) = sRecord(i) & Trim(rs("VslCode")) & ","
				sRecord(i) = sRecord(i) & sVoyage & ","
				sRecord(i) = sRecord(i) & Trim(rs("ContNo")) & ","
				sRecord(i) = sRecord(i) & Trim(rs("BLNo")) & ","
				sRecord(i) = sRecord(i) & sDateTime1

				rs.MoveNext
				i = i+1
			Loop
		End If
		rs.Close

	Case "C":
		sOpeName = "��q�ɒ�"
		sSQL = "SELECT ExportCont.VslCode, ExportCont.ContNo, ExportCont.BookNo, " & _
			" Container.ShipLine, VslSchedule.LdVoyage " & _
			" FROM ExportCont, Container, VslSchedule " & _
			" WHERE ExportCont.ContNo = '" & sContNo & "' " & _
			" AND Container.VslCode=*ExportCont.VslCode " & _
			" AND Container.VoyCtrl=*ExportCont.VoyCtrl " & _
			" AND Container.ContNo=*ExportCont.ContNo " & _
			" AND VslSchedule.VslCode=*ExportCont.VslCode " & _
			" AND VslSchedule.VoyCtrl=*ExportCont.VoyCtrl "
		rs.Open sSQL, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "�Ώۺ��łȂ�"
		Else
			Do While Not rs.Eof
				ReDim Preserve sRecord(i)

				If IsNull(rs("ShipLine")) Then
					sShipLine = ""
				Else
					sShipLine = Trim(rs("ShipLine"))
				End If
				If IsNull(rs("LdVoyage")) Then
					sVoyage = ""
				Else
					sVoyage = Trim(rs("LdVoyage"))
				End If
				sRecord(i) = sRecord(i) & sDailyTransNo & ","
				sRecord(i) = sRecord(i) & "EX04,"
				sRecord(i) = sRecord(i) & "R,"
				sRecord(i) = sRecord(i) & sDateTime2 & ","
				sRecord(i) = sRecord(i) & sUserID & ","
				sRecord(i) = sRecord(i) & TRANS_PLACE & ","
				sRecord(i) = sRecord(i) & Trim(rs("VslCode")) & ","
				sRecord(i) = sRecord(i) & sVoyage & ","
				sRecord(i) = sRecord(i) & Trim(rs("ContNo")) & ","
				sRecord(i) = sRecord(i) & Trim(rs("BookNo")) & ","
				sRecord(i) = sRecord(i) & sShipLine & ","
				sRecord(i) = sRecord(i) & sDateTime1

				rs.MoveNext
				i = i+1
			Loop
		End If
		rs.Close

	' ����ݸފ�
	Case "D":
		sOpeName = "����ݸފ�"
		sSQL = "SELECT ExportCont.VslCode, ExportCont.ContNo, ExportCont.BookNo, " & _
			" Container.ShipLine, VslSchedule.LdVoyage " & _
			" FROM ExportCont, Container, VslSchedule " & _
			" WHERE ExportCont.ContNo = '" & sContNo & "' " & _
			" AND Container.VslCode=*ExportCont.VslCode " & _
			" AND Container.VoyCtrl=*ExportCont.VoyCtrl " & _
			" AND Container.ContNo=*ExportCont.ContNo " & _
			" AND VslSchedule.VslCode=*ExportCont.VslCode " & _
			" AND VslSchedule.VoyCtrl=*ExportCont.VoyCtrl "
		rs.Open sSQL, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "�Ώۺ��łȂ�"
		Else
			Do While Not rs.Eof
				ReDim Preserve sRecord(i)

				If IsNull(rs("ShipLine")) Then
					sShipLine = ""
				Else
					sShipLine = Trim(rs("ShipLine"))
				End If
				If IsNull(rs("LdVoyage")) Then
					sVoyage = ""
				Else
					sVoyage = Trim(rs("LdVoyage"))
				End If
				sRecord(i) = sRecord(i) & sDailyTransNo & ","
				sRecord(i) = sRecord(i) & "EX05,"
				sRecord(i) = sRecord(i) & "R,"
				sRecord(i) = sRecord(i) & sDateTime2 & ","
				sRecord(i) = sRecord(i) & sUserID & ","
				sRecord(i) = sRecord(i) & TRANS_PLACE & ","
				sRecord(i) = sRecord(i) & Trim(rs("VslCode")) & ","
				sRecord(i) = sRecord(i) & sVoyage & ","
				sRecord(i) = sRecord(i) & Trim(rs("ContNo")) & ","
				sRecord(i) = sRecord(i) & Trim(rs("BookNo")) & ","
				sRecord(i) = sRecord(i) & sShipLine & ","
				sRecord(i) = sRecord(i) & sDateTime1 & ","
				sRecord(i) = sRecord(i) & ","

				rs.MoveNext
				i = i+1
			Loop
		End If
		rs.Close
End Select

conn.Close

Set oFS = Server.CreateObject("Scripting.FileSystemObject")

If sRecord(0) <> "" Then
	Dim sTransPath
	Dim sFileName
	Dim oFS
	Dim oTS

	' �g�����U�N�V�����t�@�C���쐬
	sFileName = ArrangeNum(Month(tNow), 2) & ArrangeNum(Day(tNow), 2) & sDailyTransNo
	sTransPath = Server.MapPath(SEND_FOLDER & "/" & sFileName & "." & SEND_EXTENT)

	Set oTS=oFS.OpenTextFile(sTransPath, 2, True)
	For j=0 To i-1
	    oTS.WriteLine sRecord(j)
	Next
	oTS.Close
	Set oTS = Nothing

End If

' Log�o��
If sErrMsg<>"" Then
	WriteLogM oFS, sUserID, "6202", "�g��-������������", "10",sPhoneType, sContNo & "/" & sOpeName & "," & "���͓��e�̐���:1(���)" & sErrMsg
Else
	WriteLogM oFS, sUserID, "6202", "�g��-������������", "10",sPhoneType, sContNo & "/" & sOpeName & "," & "���͓��e�̐���:0(������)"
End If

Set oFS = Nothing

Response.Buffer = TRUE
Response.Expires = 0
If sErrMsg = "" Then
	Response.Redirect Application("URL_MOBILE") & "index.asp"
	Response.End
End If

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="������������">
		<center>
		�y�����������́z<br><br>
		<center>
		<%=sErrMsg%><br>
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
		<%=GetTitleTag("������������")%>
	</head>
	<body>
	<center>
	�y�����������́z
	<hr>
	<%=sErrMsg%><br>
	<br>
	<form action="index.asp" method="get">
		<input type="submit" value="�ƭ�">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
