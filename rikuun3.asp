<%@ LANGUAGE="VBScript" %>

<!--#include file="common.inc"-->
<!--#include file="ija/mcommon.inc"-->
<%
    ' セッションのチェック
    CheckLogin "rikunn1.asp"

    ' ユーザ種類を取得する
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' セッションが切れているとき
        Response.Redirect "index.asp"             'トップ
        Response.End
    End If

' トランザクションファイルを出力するフォルダ
Const SEND_FOLDER = "send"

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

sUserID = Trim(Session.Contents("userid"))
sContNo = Trim(Session.Contents("cntnrno"))
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

	Dim sTransPath
	Dim sFileName
	Dim oFS
	Dim oTS

Set oFS = Server.CreateObject("Scripting.FileSystemObject")

If sRecord(0) <> "" Then

	' トランザクションファイル作成
	sFileName = ArrangeNum(Month(tNow), 2) & ArrangeNum(Day(tNow), 2) & sDailyTransNo
	sTransPath = Server.MapPath(SEND_FOLDER & "\" & sFileName & "." & SEND_EXTENT)

	Set oTS=oFS.OpenTextFile(sTransPath, 2, True)
	For j=0 To i-1
	    oTS.WriteLine sRecord(j)
	Next
	oTS.Close
	Set oTS = Nothing

End If

' Log出力
If sErrMsg<>"" Then
	WriteLog oFS, "6002", "陸運入力-完了時刻入力", "10", sContNo & "/" & sOpeName & "," & "入力内容の正誤:1(誤り)" & sErrMsg
Else
	WriteLog oFS, "6002", "陸運入力-完了時刻入力", "10", sContNo & "/" & sOpeName & "," & "入力内容の正誤:0(正しい)"
End If
Set oFS = Nothing


Response.Buffer = TRUE
Response.Expires = 0
If sErrMsg = "" Then
	Response.Redirect "http://www.hits-h.com/index.asp"
	Response.End
End If
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="gif/rikuunt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
    <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
<center>
<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>運行情報入力</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
		<br><br>
<%
    DispErrorMessage sErrMsg
%>
		</center></td>
 </tr>
 <tr>
    <td valign="bottom">
<%
    DispMenuBar
%>
    </td>
  </tr>
</table>
 </td>
 </tr>
 </table>
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "rikuun1.asp"
%>
</body>
</html>

