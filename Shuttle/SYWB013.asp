<%@ LANGUAGE="VBScript" %>
<%

' Added by Seiko-denki 2003.07.24
'認証方法変更のためコメントアウト20040204 S
'	if request.querystring("UserId")<>"" then
'		strInputUserID = request.querystring("UserId")
'	else
'		if Session.Contents("userid") <> "" then
'			strInputUserID = Session.Contents("userid")
'		else
'			strInputUserID = ""
'		end if
'	end if
'
'	if strInputUserID<>"" and Session.Contents("login_count")=1 then
'		Session.Contents("userid")=strInputUserID
'	else
'		Session.Contents("login_count")=1
'		response.redirect "http://www.cont-info.com/Userchk2.asp?ReturnUrl=http://www.hits-h.com/SYWB013.asp"
'	end if
'認証方法変更のためコメントアウト20040204 E
' End of Addition by Seiko-denki 2003.07.24


'Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB013.inc"-->
<!--#include file="SYWB077.inc"-->
<html>

<head>
<title>搬出入予約申請作業一覧画面</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
function SelDate(sel) {
//  location.href = "SYWB013.asp?TDATE=" + sel.options[sel.selectedIndex].value;
//	location.reload(true);
	location.replace("SYWB013.asp?UserId=<%=strInputUserID%>&TDATE=" + sel.options[sel.selectedIndex].value);
}
</SCRIPT>
</head>

<body>
<%
	'VP対応2001/8/22
	'**********   ＤＢ接続情報   **********
	Dim conn, rsd			'ＤＢ接続
	'**********   日付情報   **********
	Dim sTrgDate			'指定日付("YYYYMMDD")
	Dim sDateNow			'現在日付("YYYYMMDD")
	'**********   ユーザ情報   **********
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator	'ユーザ情報
	'**********   運行状況情報   **********
	Dim iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime, iNextApp	'運行状況
	'**********   運行時間情報   **********
	Dim sShtStart, sShtEnd	'シャトル運行時間（HHMM）
	Dim iSTime, iETime		'シャトル運行時間（時間帯）
	Dim iTimeCnt			'時間帯数
	Dim TimeSlot(30)		'時間帯記号（インデックス＝行番号）例：08～16,A,B,D
	Dim TimeNo(30)			'時間帯番号（インデックス＝行番号）例：8～16,30,31,32
	Dim iRecDelCnt(30, 1)	'搬出入本数（インデックス＝行番号）
 	Dim sOpenFlag(23)		'開放フラグ（インデックス＝時間帯番号）
	Dim iCloseMode(30)		'完了モード（インデックス＝行番号）
							'（0：運行前　1：完了　2：運行中　3：確定　4：確定中　-1：開放中）
	Dim TimeName(30), TimeJmp(30), sStatus(30)		'時間帯表示情報
	Dim iLuckChassis(30, 1), sLuckChassis(30, 1)	'予測空シャーシ数
	'**********   データ制御   **********
	Dim iLineCnt(30)				'時間帯ごとの表示行数
	Dim iRecIdx(30, 100)			'表示行と申請情報の対応テーブル
									'  0～n：インデックス　-1：なし　-2：搬入なし　-3：搬出なし
	'**********   その他   **********
	Dim iEmptySlot, iEmptyChassis(1)			'空きスロット、空シャーシ
	Dim iCnt, iWk, sWk, bWk, i, k, sColor(9)
	Dim sDate
	Dim sDays(20), iDaysCnt						'営業日
	Dim sCell(16)
	'**********   申請情報（インデックス＝レコード）   **********
	Dim iAppCnt					'申請情報数
	Dim iAppOpeNo(1000)			'作業番号
	Dim sAppUserNm(1000)		'ユーザ名
	Dim sAppContNo(1000)		'コンテナ番号
	Dim sAppBLNo(1000)			'ＢＬ番号
	Dim sAppRecDel(1000)		'搬出入区分
	Dim sAppStatus(1000)		'状態
	Dim sAppPlace(1000)			'場所
    Dim sAppChassisId(1000)		'シャーシID
	Dim sAppWorkFlag(1000)		'作業中フラグ
	Dim sAppCReason(1000)		'キャンセル理由
	Dim sAppTerm(1000)			'時間帯
	Dim sAppHopeTerm(1000)		'希望時間帯
	Dim iAppOpeOrder(1000)		'作業順位
	Dim iAppDualOpeNo(1000)		'デュアル作業番号
	Dim sAppContSize(1000)		'コンテナサイズ
	Dim sAppFromTo(1000)		'搬出先／搬入元
	Dim sAppDelFlag(1000)		'削除フラグ
	Dim sDelChaStock(1000)		'搬出指定シャーシの在庫
	Dim sAppTerminal(1000)		'ターミナルコード		'VP対応(01/10/01)
	Dim sAppVPBookNo(1000)		'ＶＰブッキング番号		'VP対応(01/10/01)

	'指定日付取得
	sTrgDate = TRIM(Request.QueryString("TDATE"))

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	If sGrpID = "" Then
		Response.Write "ユーザが登録されていません。(" & sUsrID & ")"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.End
	End If

	'空きスロットの取得
	Call GetLackChassis(conn, rsd, sGrpID, _
			iEmptySlot, iEmptyChassis(0), iEmptyChassis(1))

	'現在日付取得
	sDateNow = GetYMDStr(Date())

	'営業日の取得
	Call GetBusinessDays(conn, rsd, sDateNow, iDaysCnt, sDays)

	'指定日付がない場合はデフォルトをセット
	If sTrgDate = "" Then
		sTrgDate = Trim(sDays(1))
	End If

	'運行状況を取得
	Call GetOpeStatusDtl(conn, rsd, _
						iCurTime, iNextTime, iNextStat, _
						iOpenSlot, sEndTime, iNextApp)

	'夕積終了予定を計算
	If sEndTime = "" Then
		sEndTime = "未定"
	Else
		sEndTime = Left(sEndTime, 2) & ":" & Right(sEndTime, 2)
	End If

	If sTrgDate <> "WAIT" Then	'通常の場合
		'グループ時間帯情報の取得（開放の有無を取得）
		Call GetGrpSlot(conn, rsd, sGrpID, sTrgDate, sOpenFlag)

		'シャトル運行時間取得
		sShtStart = GetEnv(conn, rsd, "ShtStart")
		sShtEnd   = GetEnv(conn, rsd, "ShtEnd")
		iSTime = CLng(Left(sShtStart, 2))
		iETime = CLng(Left(sShtEnd, 2))
		If Right(sShtEnd, 2) = "00" Then
			iETime = iETime - 1
		End If

		'シャトル運行時間帯計算
		iCnt = 0

		'時間帯の計算
		''午前時間
		For i = iSTime To 12
			TimeSlot(iCnt) = Right("0" & CStr(i), 2)
			TimeNo(iCnt) = i
			iCnt = iCnt + 1
		Next
		''午後時間
		For i = 13 To iETime
			TimeSlot(iCnt) = Right("0" & CStr(i), 2)
			TimeNo(iCnt) = i
			iCnt = iCnt + 1
		Next
		''午後指定
		TimeSlot(iCnt) = "A"
		TimeNo(iCnt) = 30
		iCnt = iCnt + 1
		''夕積指定
		TimeSlot(iCnt) = "B"
		TimeNo(iCnt) = 31
		iCnt = iCnt + 1
		''ユーザ削除
		TimeSlot(iCnt) = "D"
		TimeNo(iCnt) = 32
		iCnt = iCnt + 1

		iTimeCnt = iCnt		'時間帯数

		'申請情報取得
		iAppCnt = 0
		For i = 0 To iTimeCnt - 1
			'指定時間帯、指定グループの申請情報を取得
			Call GetAppHH(conn, rsd, _
					sGrpID, sTrgDate, TimeSlot(i), TimeNo(i), _
					sDateNow, iCurTime, iNextTime, iNextApp, _
					iRecDelCnt(i, 0), iRecDelCnt(i, 1), iCloseMode(i), _
					iAppCnt, _
					iAppOpeNo, sAppUserNm, sAppContNo, _
					sAppBLNo, sAppRecDel, sAppStatus, _
					sAppPlace, sAppChassisId, _
					sAppWorkFlag, sAppCReason, sAppContSize, _
					sAppTerm, sAppHopeTerm, iAppOpeOrder, _
					iAppDualOpeNo, sAppFromTo, sAppDelFlag, sDelChaStock, sAppTerminal, sAppVPBookNo)
		Next

		'シャーシ設定
		''デュアルで搬入側シャーシが決定している場合に搬出側にシャーシをセット
		Call SetAppChas(iAppCnt, _
						iAppOpeNo, sAppUserNm, sAppContNo, _
						sAppBLNo, sAppRecDel, sAppStatus, _
						sAppPlace, sAppChassisId, _
						sAppWorkFlag, sAppCReason, sAppContSize, _
						sAppTerm, sAppHopeTerm, iAppOpeOrder, _
						iAppDualOpeNo, sAppFromTo)

		'本日が運行日の場合は仮シャーシ計算を行う
		If sTrgDate = sDateNow Then
			'仮シャーシ計算
			Call CalcAppChas(conn, rsd, _
						sGrpID, sTrgDate, _
						iCurTime, iNextTime, iNextStat, _
						iAppCnt, _
						iAppOpeNo, sAppUserNm, sAppContNo, _
						sAppBLNo, sAppRecDel, sAppStatus, _
						sAppPlace, sAppChassisId, _
						sAppWorkFlag, sAppCReason, sAppContSize, _
						sAppTerm, sAppHopeTerm, iAppOpeOrder, _
						iAppDualOpeNo, sAppFromTo)
		End If

		'時間帯セルの設定
		For i = 0 To iTimeCnt - 1
			Call SetCell01( conn, rsd, sTrgDate, TimeSlot(i), _
							sShtStart, sShtEnd, iSTime, iETime, _
							iCloseMode(i), sOpenFlag, TimeName(i), TimeJmp(i), sStatus(i))
		Next
	Else		'引取り待ち表示の場合
		'指定グループの引取り待ち申請を取得
		Call GetAppWait(conn, rsd, _
					sGrpID, _
					iAppCnt, _
					iAppOpeNo, sAppUserNm, sAppContNo, _
					sAppBLNo, sAppRecDel, sAppStatus, _
					sAppPlace, sAppChassisId, _
					sAppWorkFlag, sAppCReason, sAppContSize, _
					sAppTerm, sAppHopeTerm, iAppOpeOrder, _
					iAppDualOpeNo, sAppFromTo, sAppDelFlag, sAppTerminal, sAppVPBookNo)
	End If
%>

<IMG border=0 height=42 src="image/title01.gif" width=311>
<br><br>
<center>
<p><IMG border=0 height=34 src="image/title21.gif" width="236" height="34"><p>
</center>
<center>
	<b>◆<% response.write sGrpName %>グループの情報です◆<b>
</center><br>
<center>
          <TD align=middle height="36">
			<A href="SYWB017.asp?YMD=<%=sTrgDate%>">シャトル利用回数</A>　　　
            <A href="../index.asp">メニューへ</A> 
          </TD>
</center><br>

<%
	rsd.Open "sUseDB", conn, 0, 1, 2
%>
<center>
<b>現在の在庫情報は　<U><%=Month(rsd("OutUpdtTime" & rsd("EnableDB")))%>月<%=Day(rsd("OutUpdtTime" & rsd("EnableDB")))%>日　
						<%=FormatDateTime(rsd("OutUpdtTime" & rsd("EnableDB")), vbShortTime)%></U>　のものです　
（<%=FormatDateTime(rsd("OutPUpdtTime"), vbShortTime)%>に更新予定)。
</b>

<%
	rsd.Close
%>
</center>

<%
	If sTrgDate <> "WAIT" Then	'通常表示の場合
		'行情報をセット
		For iCnt = 0 To iTimeCnt - 1		'時間帯数
			iLineCnt(iCnt) = 0		'時間帯ごとの表示行数
			bWk = False		'デュアルフラグ
			For i = 0 To iAppCnt - 1		'申請情報数
				If sAppTerm(i) = TimeSlot(iCnt) Then	'時間帯が一致
					If TimeSlot(iCnt) <> "12" and _
					   TimeSlot(iCnt) <> "A" and _
					   TimeSlot(iCnt) <> "B" and _
					   TimeSlot(iCnt) <> "D" Then		'時間指定の場合
						If sAppRecDel(i) = "R" Then			'搬入の場合

							'データ行追加
							iRecIdx(iCnt, iLineCnt(iCnt)) = i	'表示行と申請情報の対応テーブル
							iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'時間帯ごとの表示行数

							If iAppDualOpeNo(i) > 0 Then		'デュアルの場合
								bWk = True	'デュアルフラグ＝オン
							ElseIf sAppStatus(i) <> "03" Then	'単独でキャンセルでない場合
								'空行追加
								iRecIdx(iCnt, iLineCnt(iCnt)) = -3	'表示行と申請情報の対応テーブル
								iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'時間帯ごとの表示行数
							End If
						Else								'搬出の場合
							If bWk Then							'デュアルの場合
								bWk = False		'デュアルフラグ＝オフ
							ElseIf sAppStatus(i) <> "03" Then	'単独でキャンセルでない場合
								'空行追加
								iRecIdx(iCnt, iLineCnt(iCnt)) = -2	'表示行と申請情報の対応テーブル
								iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'時間帯ごとの表示行数
							End If
							'データ行追加
							iRecIdx(iCnt, iLineCnt(iCnt)) = i	'表示行と申請情報の対応テーブル
							iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'時間帯ごとの表示行数
						End If
					Else								'時間指定以外の場合
						iRecIdx(iCnt, iLineCnt(iCnt)) = i		'表示行と申請情報の対応テーブル
						iLineCnt(iCnt) = iLineCnt(iCnt) + 1		'時間帯ごとの表示行数
					End If
				End If
			Next
			'時間帯の表示行数が０の場合は空行を追加
			If iLineCnt(iCnt) = 0 Then	'時間帯ごとの表示行数
				iRecIdx(iCnt, iLineCnt(iCnt)) = -1	'表示行と申請情報の対応テーブル
				iLineCnt(iCnt) = 1					'時間帯ごとの表示行数
			End If
		Next

		'不足シャーシ数の計算
		For iCnt = 0 To iTimeCnt - 1
			sLuckChassis(iCnt, 0) = "-"
			sLuckChassis(iCnt, 1) = "-"
			If iCloseMode(iCnt) <> 1 And _
			   TimeSlot(iCnt) <> "D" Then	'完了以外＆削除以外
				'作業の有無判定
				bWk = False
				For i = 0 To iLineCnt(iCnt) - 1
					iWk = iRecIdx(iCnt, i)
					If iWk > -1 Then	'空白行でない
						If sAppStatus(iWk) = "02" and _
						   sAppWorkFlag(iWk) <> "Y" and _
						   iAppDualOpeNo(iWk) = 0 Then

							bWk = True
							Exit For
						End If
					End If
				Next
				If bWk Then		'作業がある場合のみ
					'空きシャーシ数の取得working
					Call GetEmptyChassisCnt(conn, rsd, _
										sGrpID, _
										sTrgDate, _
										TimeSlot(iCnt), _
										iEmptyChassis(0), iEmptyChassis(1))
					sLuckChassis(iCnt, 0) = CStr(iEmptyChassis(0))
					sLuckChassis(iCnt, 1) = CStr(iEmptyChassis(1))
				End If
			End If
		Next
	Else		'引取り待ち表示の場合
		'行情報をセット
		iTimeCnt = 1
		TimeSlot(0) = "X"
		TimeNo(0) = 0
		TimeName(0) = "　"
		TimeJmp(0) = ""
		sStatus(0) = ""
		iCloseMode(0) = 1
		iRecDelCnt(0, 0) = "-"
		iRecDelCnt(0, 1) = "-"
		sLuckChassis(0, 0) = "-"
		sLuckChassis(0, 1) = "-"

		iLineCnt(0) = iAppCnt
		For i = 0 To iAppCnt - 1
			iRecIdx(0, i) = i
		Next
	End If

	'ＤＢ切断
	conn.Close
%>
<br>
<center>
<table border="1">   

	<tr ALIGN=middle>
<td BGCOLOR=#F08080><select id=selectdate name=selectdate onChange="SelDate(this)">
<%
	'営業日メニュー作成
	For iCnt = 0 To iDaysCnt - 1
		sWk = ""
		If sTrgDate = sDays(iCnt) Then
			sWk = "SELECTED"
		End If
%>
	<option <%=sWk%> VALUE ="<%=sDays(iCnt)%>"><%=ChgYMDStr3(sDays(iCnt))%></option>
<%
	Next

	sWk = ""
	If sTrgDate = "WAIT" Then
		sWk = "SELECTED"
	End If
%>
	<option <%=sWk%> VALUE = "WAIT">引取り待ち</option>
</select></td>
		<td width="120" bgcolor ="#000080"><FONT COLOR="#ffffff">開放残枠（本）</FONT></td>
		<td width="50" BGCOLOR=#F08080><%=CStr(iOpenSlot)%></td>
		<td width="120" bgcolor ="#000080"><FONT COLOR="#ffffff">空スロット（本）</FONT></td>
		<td width="50" BGCOLOR=#F08080><%=iEmptySlot%></td>
	</tr>
</table>
</center>
<br>
		<font face="ＭＳ ゴシック">
		<center>
		<table border="1" width="930"  bgcolor = "#ffffff">   
			<tr ALIGN=middle bgcolor="#e8ffe8">
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>作業時間　</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>全本数<br>入/出</TH>
			    <TH BGCOLOR=#7FFFD4 COLSPAN=2>空ｼｬｰｼ<br>過不足</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>予約<br>番号</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2 width="20">順番</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>コンテナ／ＢＬ<br>／ブッキング</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>予約<br>タイプ</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2 width="20">種類</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2 width="20">サイズ</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2><A href="SYWB021.asp?YMD=<%=sTrgDate%>">ｼｬｰｼID</A></TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>対象<br>ＣＹ／ＶＰ</TH>		<!--空バン対応 -->
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>場所</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>状態</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>備考</TH>
			</tr>
			<tr ALIGN=middle bgcolor="#e8ffe8">
			    <TH BGCOLOR=#7FFFD4>20</TH>
			    <TH BGCOLOR=#7FFFD4>40</TH>
			</tr>
<%
	For iCnt = 0 To iTimeCnt - 1	'時間帯数
		iWk = iLineCnt(iCnt)			'時間帯ごとの表示行数
		If iWk > 0 Then
			'データセルの設定(01/10/02 VP対応)
			Call SetCell05(iRecIdx(iCnt, 0), iCloseMode(iCnt), _
					sTrgDate, TimeSlot(iCnt), TimeName(iCnt), _
					iAppOpeNo,  sAppUserNm, sAppContNo, sAppBLNo, _
          			sAppRecDel, sAppStatus,  sAppPlace, _
          			sAppChassisId, sAppWorkFlag, sAppCReason, _
          			sAppContSize, sAppTerm, sAppHopeTerm, _
					iAppOpeOrder, iAppDualOpeNo, sAppFromTo, _
					sAppDelFlag, sDelChaStock, sAppTerminal, sAppVPBookNo, _
					sCell)
			'時間帯セルカラーの計算
			sColor(0) = ""
			sColor(1) = ""
			sColor(3) = "bgcolor=""#AFEEEE"" "
			If TimeSlot(iCnt) = "D" Then
				sColor(0) = "bgcolor=""#dda0dd"" "
				sColor(3) = "bgcolor=""#dda0dd"" "
			End If
			If TimeSlot(iCnt) = "12" Or _
			   TimeSlot(iCnt) = "A" Or _
			   TimeSlot(iCnt) = "B" Then
				sColor(1) = "bgcolor=""#FFD700"" "
			Else
				If TimeSlot(iCnt) = "D" Then
					sColor(1) = "bgcolor=""#dda0dd"" "
				Else
					If sStatus(iCnt) = "完了" Then
						sColor(1) = "bgcolor=""#c0c0c0"" "
					ElseIf sStatus(iCnt) = "運行中" Then
						sColor(1) = "bgcolor=""#F08080"" "
					Else
						sColor(1) = "bgcolor=""#FFFFE0"" "
					End If
				End If
			End If
			'データセルカラーの計算
			Call CalcDataColor(sColor(2), sCell)
%>
			<tr ALIGN=middle <%=sColor(0)%>>
<% 'ユーザ削除センタ表示(2001/03/23)
				If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
			    <td <%=sColor(1)%> ROWSPAN=2>
							<%=TimeJmp(iCnt) & TimeName(iCnt)%></A>
			    <td <%=sColor(3)%> ROWSPAN=2><%=iRecDelCnt(iCnt, 0)%>/<%=iRecDelCnt(iCnt, 1)%><br></td>
<%				Else	
					If TimeName(iCnt) =  "ユーザ削除" Then	%>
				    <td <%=sColor(1)%> ROWSPAN=<%=iWk%>>
								<%=TimeJmp(iCnt) & TimeName(iCnt)%></A>
						<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=iRecDelCnt(iCnt, 0)%>/<%=iRecDelCnt(iCnt, 1)%></td>
<%					Else	%>
						<td <%=sColor(1)%> ROWSPAN=<%=iWk%>>
									<%=TimeJmp(iCnt) & TimeName(iCnt)%></A><br>
									<%=sStatus(iCnt)%></A></td>
						<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=iRecDelCnt(iCnt, 0)%>/<%=iRecDelCnt(iCnt, 1)%></td>
<%					End If	%>
<%				End If	%>
<% '空シャーシ過不足マイナス時赤字表示(2001/03/09)
				If sLuckChassis(iCnt, 0) <> "-" Then
					If sLuckChassis(iCnt, 0) < 0 Then	
						If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
					    <td <%=sColor(3)%>  ROWSPAN=2><FONT color=Red><B><%=sLuckChassis(iCnt, 0)%></B></FONT></td>
<%						Else	%>
					    <td <%=sColor(3)%>  ROWSPAN=<%=iWk%>><FONT color=Red><B><%=sLuckChassis(iCnt, 0)%></B></FONT></td>
<%						End if	
					Else				
						If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
					    <td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 0)%></td>
<%						Else	%>
					    <td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 0)%></td>
<%						End if
					End If					
				Else						
					If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
					<td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 0)%><br></td>
<%					Else	%>
					<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 0)%></td>
<%					End If
				End If					

				If sLuckChassis(iCnt, 1) <> "-" Then
					If sLuckChassis(iCnt, 1) < 0 Then	
						If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
					    <td <%=sColor(3)%>  ROWSPAN=2><FONT color=Red><B><%=sLuckChassis(iCnt, 1)%></B></FONT></td>
<%						Else	%>
					    <td <%=sColor(3)%>  ROWSPAN=<%=iWk%>><FONT color=Red><B><%=sLuckChassis(iCnt, 1)%></B></FONT></td>
<%						End if	
					Else				
						If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
					    <td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 1)%></td>
<%						Else	%>
					    <td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 1)%></td>
<%						End if
					End If					
				Else						
					If TimeName(iCnt) =  "ユーザ削除" And iWk <  2 Then	%>
					<td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 1)%><br></td>
<%					Else	%>
					<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 1)%></td>
<%					End If
				End If	%>				

			    <td <%=sColor(2)%>><%=sCell(1)%></td>
			    <td <%=sColor(2)%>><%=sCell(2)%></td>
			    <td <%=sColor(2)%>><%=sCell(3)%></td>
			    <td <%=sColor(2)%>><%=sCell(4)%></td>
			    <td <%=sColor(2)%>><%=sCell(5)%></td>
			    <td <%=sColor(2)%>><%=sCell(6)%></td>
			    <td <%=sColor(2)%>><%=sCell(7)%></td>
			    <td <%=sColor(2)%>><%=sCell(0)%></td>			<!--空バン対応 -->
			    <td <%=sColor(2)%>><%=sCell(8)%></td>
			    <td <%=sColor(2)%>><%=sCell(9)%></td>
			    <td <%=sColor(2)%>><%=sCell(10)%></td>
			</tr>
<%
		End If
		For i = 1 To iWk - 1		'時間帯ごとの表示行数-1
			'データセルの設定
			Call SetCell05(iRecIdx(iCnt, i), iCloseMode(iCnt), _
					sTrgDate, TimeSlot(iCnt), TimeName(iCnt), _
					iAppOpeNo, sAppUserNm, sAppContNo, sAppBLNo, _
          			sAppRecDel, sAppStatus, sAppPlace, _
          			sAppChassisId, sAppWorkFlag, sAppCReason, _
          			sAppContSize, sAppTerm, sAppHopeTerm, _
					iAppOpeOrder, iAppDualOpeNo, sAppFromTo, _
					sAppDelFlag, sDelChaStock, sAppTerminal, sAppVPBookNo, _
					sCell)
			'データセルカラーの計算
			Call CalcDataColor(sColor(2), sCell)
%>

			<tr ALIGN=middle <%=sColor(0)%>>
			    <td <%=sColor(2)%>><%=sCell(1)%></td>
			    <td <%=sColor(2)%>><%=sCell(2)%></td>
			    <td <%=sColor(2)%>><%=sCell(3)%></td>
			    <td <%=sColor(2)%>><%=sCell(4)%></td>
			    <td <%=sColor(2)%>><%=sCell(5)%></td>
			    <td <%=sColor(2)%>><%=sCell(6)%></td>
			    <td <%=sColor(2)%>><%=sCell(7)%></td>
			    <td <%=sColor(2)%>><%=sCell(0)%></td>			<!--空バン対応 -->
			    <td <%=sColor(2)%>><%=sCell(8)%></td>
			    <td <%=sColor(2)%>><%=sCell(9)%></td>
			    <td <%=sColor(2)%>><%=sCell(10)%></td>
			</tr>
<%
		Next
	Next
%>
		</table>
		</center>
		</font>

<br>     
<br>     
</body>     
</html>     
