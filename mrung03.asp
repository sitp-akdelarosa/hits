<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
' トランザクションファイルを出力するフォルダ
Const SEND_FOLDER = "../send"

' トランザクションファイルの拡張子
Const SEND_EXTENT = "snd"

' 発生場所
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
		sOpeName = "実入倉庫着"
		sSQL = "SELECT ImportCont.VslCode, ImportCont.ContNo, ImportCont.BLNo, VslSchedule.DsVoyage " & _
			" FROM ImportCont, VslSchedule " & _
			" WHERE ImportCont.ContNo = '" & sContNo & "' " & _
			" AND VslSchedule.VslCode=*ImportCont.VslCode " & _
			" AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl "
		rs.Open sSQL, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "対象ｺﾝﾃﾅなし"
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
		sOpeName = "デバン完"
		sSQL = "SELECT ImportCont.VslCode, ImportCont.ContNo, ImportCont.BLNo, VslSchedule.DsVoyage " & _
			" FROM ImportCont, VslSchedule " & _
			" WHERE ImportCont.ContNo = '" & sContNo & "' " & _
			" AND VslSchedule.VslCode=*ImportCont.VslCode " & _
			" AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl "
		rs.Open sSQL, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "対象ｺﾝﾃﾅなし"
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
		sOpeName = "空倉庫着"
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
			sErrMsg = "対象ｺﾝﾃﾅなし"
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

	' ﾊﾞﾝﾆﾝｸﾞ完
	Case "D":
		sOpeName = "ﾊﾞﾝﾆﾝｸﾞ完"
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
			sErrMsg = "対象ｺﾝﾃﾅなし"
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

	' トランザクションファイル作成
	sFileName = ArrangeNum(Month(tNow), 2) & ArrangeNum(Day(tNow), 2) & sDailyTransNo
	sTransPath = Server.MapPath(SEND_FOLDER & "/" & sFileName & "." & SEND_EXTENT)

	Set oFS = Server.CreateObject("Scripting.FileSystemObject")
	Set oTS=oFS.OpenTextFile(sTransPath, 2, True)
	    oTS.WriteLine sRecord
	oTS.Close
	Set oTS = Nothing

	' Log出力
	WriteLogM oFS, sUserID, "運行情報入力(携帯)", sOpeName & "(" & sContNo & ")"

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
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="運行情報入力">
		<center>
		<img src="mtitle<%=GetImageExt()%>" alt="一貫物流システム"><br>
		<center>
		【運行情報入力】<br><br>
		<center>
		<%=sErrMsg%><br>
		<center>
		<a task="gosub" dest="http://www.hits-h.com/index.asp">ﾒﾆｭｰ</a>
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
		<%=GetTitleTag("運行情報入力")%>
	</head>
	<body>
	<center>
	<img src="mtitle<%=GetImageExt()%>" alt="一貫物流システム"><br>
	【運行情報入力】
	<hr>
	<%=sErrMsg%><br>
	<br>
	<form action="http://www.hits-h.com/index.asp" method="get">
		<input type="submit" value="ﾒﾆｭｰ">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
