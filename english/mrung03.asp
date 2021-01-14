<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
' �g�����U�N�V�����t�@�C�����o�͂���t�H���_
Const SEND_FOLDER = "../send"

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
Dim sRecord
Dim sDateTime1, sDateTime2
Dim conn, rs
Dim sSQL
Dim sVoyage, sShipLine
Dim sErrMsg

sErrMsg = ""

sUserID = Trim(Request.QueryString("UserID"))
sContNo = Trim(Request.QueryString("ContNo"))
sOperation = Trim(Request.QueryString("operation"))

sDailyTransNo = GetDailyTransNo()
tNow = Now
sDateTime1 = ArrangeNum(Year(tNow), 4) & ArrangeNum(Month(tNow), 2) & ArrangeNum(Day(tNow), 2) & _
			ArrangeNum(Hour(tNow), 2) & ArrangeNum(Minute(tNow), 2)
sDateTime2 = sDateTime1 & ArrangeNum(Second(tNow), 2)

ConnectSvr conn, rs

sRecord = ""
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
			If IsNull(rs("DsVoyage")) Then
				sVoyage = ""
			Else
				sVoyage = Trim(rs("DsVoyage"))
			End If
			sRecord = sRecord & sDailyTransNo & ","
			sRecord = sRecord & "IM10,"
			sRecord = sRecord & "R,"
			sRecord = sRecord & sDateTime2 & ","
			sRecord = sRecord & sUserID & ","
			sRecord = sRecord & TRANS_PLACE & ","
			sRecord = sRecord & Trim(rs("VslCode")) & ","
			sRecord = sRecord & sVoyage & ","
			sRecord = sRecord & Trim(rs("ContNo")) & ","
			sRecord = sRecord & Trim(rs("BLNo")) & ","
			sRecord = sRecord & sDateTime1
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
			If IsNull(rs("DsVoyage")) Then
				sVoyage = ""
			Else
				sVoyage = Trim(rs("DsVoyage"))
			End If
			sRecord = sRecord & sDailyTransNo & ","
			sRecord = sRecord & "IM11,"
			sRecord = sRecord & "R,"
			sRecord = sRecord & sDateTime2 & ","
			sRecord = sRecord & sUserID & ","
			sRecord = sRecord & TRANS_PLACE & ","
			sRecord = sRecord & Trim(rs("VslCode")) & ","
			sRecord = sRecord & sVoyage & ","
			sRecord = sRecord & Trim(rs("ContNo")) & ","
			sRecord = sRecord & Trim(rs("BLNo")) & ","
			sRecord = sRecord & sDateTime1
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
			sRecord = sRecord & sDailyTransNo & ","
			sRecord = sRecord & "EX04,"
			sRecord = sRecord & "R,"
			sRecord = sRecord & sDateTime2 & ","
			sRecord = sRecord & sUserID & ","
			sRecord = sRecord & TRANS_PLACE & ","
			sRecord = sRecord & Trim(rs("VslCode")) & ","
			sRecord = sRecord & sVoyage & ","
			sRecord = sRecord & Trim(rs("ContNo")) & ","
			sRecord = sRecord & Trim(rs("BookNo")) & ","
			sRecord = sRecord & sShipLine & ","
			sRecord = sRecord & sDateTime1
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
			sRecord = sRecord & sDailyTransNo & ","
			sRecord = sRecord & "EX05,"
			sRecord = sRecord & "R,"
			sRecord = sRecord & sDateTime2 & ","
			sRecord = sRecord & sUserID & ","
			sRecord = sRecord & TRANS_PLACE & ","
			sRecord = sRecord & Trim(rs("VslCode")) & ","
			sRecord = sRecord & sVoyage & ","
			sRecord = sRecord & Trim(rs("ContNo")) & ","
			sRecord = sRecord & Trim(rs("BookNo")) & ","
			sRecord = sRecord & sShipLine & ","
			sRecord = sRecord & sDateTime1 & ","
			sRecord = sRecord & ","
		End If
		rs.Close
End Select

conn.Close

If sRecord <> "" Then
	Dim sTransPath
	Dim sFileName
	Dim oFS
	Dim oTS

	' �g�����U�N�V�����t�@�C���쐬
	sFileName = ArrangeNum(Month(tNow), 2) & ArrangeNum(Day(tNow), 2) & sDailyTransNo
	sTransPath = Server.MapPath(SEND_FOLDER & "/" & sFileName & "." & SEND_EXTENT)

	Set oFS = Server.CreateObject("Scripting.FileSystemObject")
	Set oTS=oFS.OpenTextFile(sTransPath, 2, True)
	    oTS.WriteLine sRecord
	oTS.Close
	Set oTS = Nothing

	' Log�o��
	WriteLogM oFS, sUserID, "�^�s������(�g��)", sOpeName & "(" & sContNo & ")"

	Set oFS = Nothing
End If

Response.Buffer = TRUE
Response.Expires = 0
If sErrMsg = "" Then
	Response.Redirect Application("URL_MOBILE") & "http://www.hits-h.com/index.asp"
	Response.End
End If

Dim sPhoneType
sPhoneType = GetPhoneType()
If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="�^�s������">
		<center>
		<img src="mtitle<%=GetImageExt()%>" alt="��ѕ����V�X�e��"><br>
		<center>
		�y�^�s�����́z<br><br>
		<center>
		<%=sErrMsg%><br>
		<center>
		<a task="gosub" dest="http://www.hits-h.com/index.asp">�ƭ�</a>
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
		<%=GetTitleTag("�^�s������")%>
	</head>
	<body>
	<center>
	<img src="mtitle<%=GetImageExt()%>" alt="��ѕ����V�X�e��"><br>
	�y�^�s�����́z
	<hr>
	<%=sErrMsg%><br>
	<br>
	<form action="http://www.hits-h.com/index.asp" method="get">
		<input type="submit" value="�ƭ�">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
