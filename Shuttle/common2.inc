<% 
'ＤＢ接続
Function ConnectSvr(conn, rsd) 
	'conn		[ O ]DB接続Object
	'rsd		[ O ]レコードセットObject

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "hakataDB", "sa", ""

	Set rsd = Server.CreateObject("ADODB.Recordset")

End Function

'ユーザ情報の取得
Function GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	'conn		[ I ]DB接続Object
	'rsd		[ I ]レコードセットObject
	'sUsrID		[ O ]ユーザID
	'sGrpID		[ O ]グループID
	'sUsrName	[ O ]ユーザ名
	'sGrpName	[ O ]グループ名
	'sOperator	[ O ]オペレータコード

	Dim sql
	Dim iWk

	'認証ユーザ取得
'	sUsrID = UCase(Trim(Request.ServerVariables("AUTH_USER")))  'Commented by Seiko-denki 2003.7.26
	sUsrID = Session.Contents("userid")  'Added by Seiko-denki 2003.7.26

	If sUsrID <> "" Then
		iWk	= InStrRev(sUsrID, "\")
		If iWk > 0 Then
			sUsrID = Mid(sUsrID, iWk + 1)
		End If
	End If
	sUsrID = Left(sUsrID, 8)

	sGrpID    = ""
	sUsrName  = "不明（" & sUsrID & "）"
	sGrpName  = "不明"
	sOperator = ""

	'ユーザ取得
	sql = "SELECT sMUser.Name,sMGroup.GroupName,sMGroup.GroupID,sMGroup.Operator" & _
		  " FROM sMUser,sMGroup,sMUserGroup" & _
		  " WHERE RTRIM(sMUser.UserID) = '" & sUsrID & "'" & _
		  "   AND RTRIM(sMUserGroup.UserID) = RTRIM(sMUser.UserID)" & _
		  "   AND RTRIM(sMGroup.GroupID) = RTRIM(sMUserGroup.GroupID)"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		sGrpID    = Trim(rsd("GroupID"))
		sUsrName  = Trim(rsd("Name"))
		sGrpName  = Trim(rsd("GroupName"))
		sOperator = Trim(rsd("Operator"))
	End If

	rsd.Close

End Function

'環境変数の取得
Function GetEnv(conn, rsd, sEnvName)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sEnvName	[ I ]環境変数名
'	戻り値		[ O ]環境変数値

	Dim sql

	'環境変数取得
	sql = "SELECT EnvValue FROM sEnv" & _
		  " WHERE RTRIM(EnvName) = '" & sEnvName & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		GetEnv = Trim(rsd("EnvValue"))
	Else
		GetEnv = ""
	End If

	rsd.Close

End Function

'日付文字列の取得
Function GetYMDStr(DateVal)
	'DateVal	[ I ]日付（Date型）
	'戻り値		[ O ]日付文字列（YYYYMMDD形式）

	GetYMDStr = _
		Year(DateVal) & _
		Right("0" & Month(DateVal), 2) & _
		Right("0" & Day(DateVal), 2)

End Function

'日付文字列の変換
Function ChgYMDStr(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY/MM/DD形式）

	ChgYMDStr = Left(sDate, 4) & "/" & _
				Mid(sDate, 5, 2) & "/" & _
				Mid(sDate, 7, 2)

End Function

'日付文字列の変換２
Function ChgYMDStr2(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY年MM月DD形式）

	ChgYMDStr2 = Left(sDate, 4) & "年" & _
			 	 CStr(CLng(Mid(sDate, 5, 2))) & "月" & _
				 CStr(CLng(Mid(sDate, 7, 2))) & "日"

End Function

'日付文字列の変換３
Function ChgYMDStr3(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY/MM/DD形式）

	Dim sYobi
	ChgYMDStr3 = Left(sDate, 4) & "/" & _
				Mid(sDate, 5, 2) & "/" & _
				Mid(sDate, 7, 2)
	sYobi = sWeekday(Weekday(cDate(ChgYMDStr3)))
	ChgYMDStr3 = ChgYMDStr3 & "(" & sYobi & ")"
End Function

'文字列を日付に変換
Function ChgYMDDate(sYMD)
'	sYMD		[ I ]日付文字列（YYYYMMDD形式）
'	戻り値		[ O ]日付（Date型）

	ChgYMDDate = CDate(ChgYMDStr(sYMD))
End Function

'時間を時間帯に変換
Function GetTimeSlot(iH, iM, sSE)
	'iH			[ I ]時間
	'iM			[ I ]分
	'sSE		[ I ]S:開始、E:終了
	'戻り値		[ O ]時間帯文字列

	Dim sStr

	sStr = CStr(iH) & ":" & Right("0" & CStr(iM), 2)

	If sSE = "S" Then
		sStr = sStr & "～" & CStr(iH + 1) & ":00"
	Else
		sStr = CStr(iH - 1) & ":00～" & sStr
	End If

	GetTimeSlot = sStr

End Function

'時間を時間帯文字に変換
Function GetTimeSlotStr(conn, rsd, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sHH			[ I ]時間（"HH"、"A"、"B"）
'	戻り値		[ O ]時間帯文字列

	Dim sStr, sShtStart, sShtEnd, iSTime, iETime, iHH

	Select Case Trim(sHH)
 		Case "12"
			sStr = "午前指定"
 		Case "A"
			sStr = "午後指定"
   		Case "B"
			sStr = "夕積指定"
    	Case Else 
			'開始／終了時間を取得
			sShtStart = GetEnv(conn, rsd, "ShtStart")
			sShtEnd   = GetEnv(conn, rsd, "ShtEnd")
			iSTime = CLng(Left(sShtStart, 2))
			iETime = CLng(Left(sShtEnd, 2))
			if Right(sShtEnd, 2) = "00" Then
'''2001/03/12			iETime = iETime - 1
				If iETime <> "16" Then
					iETime = iETime - 1
				End If
			End If

			iHH = CLng(sHH)
			If iHH = iSTime Then
				sStr = CStr(iHH) & ":" & Right(sShtStart, 2)
			Else
				sStr = CStr(iHH) & ":00"
			End If
			sStr = sStr & "～"
			If iHH = iETime Then
				sStr = sStr & CStr(iHH) & ":" & Right(sShtEnd, 2)
			Else
				sStr = sStr & CStr(iHH + 1) & ":00"
			End If
	End Select 

	GetTimeSlotStr = sStr

End Function

'グループ時間帯情報の取得
Function GetGrpSlot(conn, rsd, sGroupID, sDate, sOpenFlag())
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sGroupID	[ I ]グループコード
'	sDate		[ I ]日付(YYYYMMDD)
'	sOpenFlag	[ O ]開放フラグ
'	戻り値		[ O ]なし

	Dim sql
	Dim i

	'グループスロット取得
	sql = "SELECT * FROM sGrpSlot" & _
		  " WHERE RTRIM(GroupID) = '" & sGroupID & "'" & _
		  "   AND Date = '" & sDate & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		For i = 0 To 23
			sOpenFlag(i) = Trim(rsd("OpenFlag" & Trim(Cstr(i))))
		Next
	Else
		For i = 0 To 23
			sOpenFlag(i) = "Y"
		Next
	End If

	rsd.Close

End Function

'グループ時間帯情報の取得
Function GetGrpSlotHH(conn, rsd, sGroupID, sDate, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sGroupID	[ I ]グループコード
'	sDate		[ I ]日付(YYYYMMDD)
'	sHH			[ I ]時間帯（HH）
'	戻り値		[ O ]開放フラグ

	Dim sql

	'グループスロット取得
	sql = "SELECT * FROM sGrpSlot" & _
		  " WHERE RTRIM(GroupID) = '" & sGroupID & "'" & _
		  "   AND Date = '" & sDate & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		GetGrpSlotHH = Trim(rsd("OpenFlag" & CStr(CLng(sHH))))
	Else
		GetGrpSlotHH = "Y"
	End If

	rsd.Close

End Function

'作業状況の取得
Sub GetOpeStatus(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	iCurTime	[ O ]作業中の時間帯番号（8～16,31,99）
'	iNextTime	[ O ]次作業の時間帯番号（8～16,31,99）
'	iNextStat	[ O ]次作業のステータス（0：３０分より前　1：３０分以内　2：５分以内）
'	iOpenSlot	[ O ]開放枠
'	sEndTime	[ O ]夕積終了予定（HHMM）
	Dim iNextApp

	Call GetOpeStatusDtl(conn, rsd, _
						iCurTime, iNextTime, iNextStat, _
						iOpenSlot, sEndTime, iNextApp)
End Sub

'作業状況詳細の取得
Sub GetOpeStatusDtl(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime, iNextApp)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	iCurTime	[ O ]作業中の時間帯番号（8～16,31,99）
'	iNextTime	[ O ]次作業の時間帯番号（8～16,31,99）
'	iNextStat	[ O ]次作業のステータス（0：３０分より前　1：３０分以内　2：５分以内）
'	iOpenSlot	[ O ]開放枠
'	sEndTime	[ O ]夕積終了予定（HHMM）
'	iNextApp	[ O ]次作業の申請数（５分前確定時の値）

	Dim sql

	'作業状態ＤＢ読込
	sql = "SELECT * FROM sOpeStatus"
	rsd.Open sql, conn, 0, 1, 1

	iCurTime = 0
	iNextTime = 0
	If Not rsd.EOF Then		'レコードがある場合
		If IsNumeric(rsd("CurTime")) Then		'数値の場合
			iCurTime = CLng(rsd("CurTime"))
		Else
			If Trim(rsd("CurTime")) = "B" Then
				iCurTime = 31					'作業中の時間帯番号（8～16,31,99）
			Else
				iCurTime = 99					'作業中の時間帯番号（8～16,31,99）
			End If
		End If
		If IsNumeric(rsd("NextTime")) Then		'数値の場合
			iNextTime = CLng(rsd("NextTime"))
		Else
			If Trim(rsd("NextTime")) = "B" Then
				iNextTime = 31					'次作業の時間帯番号（8～16,31,99）
			Else
				iNextTime = 99					'次作業の時間帯番号（8～16,31,99）
			End If
		End If
		iNextStat = CLng(rsd("NextStatus"))		'次作業のステータス
		iOpenSlot = rsd("OpenSlot")				'開放枠
		sEndTime  = Trim(rsd("WorkEndTime"))	'夕積終了予定（HHMM）
		iNextApp  = rsd("NextApp")				'次作業の申請数（５分前確定時の値）
	End If

	rsd.Close

End Sub

'時間帯番号から時間帯記号を取得
Function TermNumToSym(iTernNum)
'	iTernNum	[ I ]時間帯番号（8～16,30,31,32,99）
'	戻り値		[ O ]時間帯記号（08～16,A,B,D,Z）

	Select Case iTernNum
   		Case 30		'午後
			TermNumToSym = "A"
    	Case 31		'夕積
			TermNumToSym = "B"
    	Case 32		'削除
			TermNumToSym = "D"
    	Case 99		'なし
			TermNumToSym = "Z"
		Case Else
			TermNumToSym = Right("0" & CStr(iTernNum), 2)
	End Select 

End Function

'時間帯記号から時間帯番号を取得
Function TermSymToNum(sTermSym)
'	sTermSym	[ I ]時間帯記号（08～16,A,B,D,Z）
'	戻り値		[ O ]時間帯番号（8～16,30,31,32,99）
	If IsNumeric(sTermSym) Then
		TermSymToNum = CLng(sTermSym)
	Else
		Select Case sTermSym
   			Case "A"	'午後
				TermSymToNum = 30
    		Case "B"	'夕積
				TermSymToNum = 31
    		Case "D"	'削除
				TermSymToNum = 32
    		Case "Z"	'なし
				TermSymToNum = 99
		End Select 
	End If
End Function

'開放枠の減算
Function SubmitOpenSlot(conn, rsd)
	'conn		[ I ]DB接続Object
	'rsd		[ I ]レコードセットObject
	'戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sOpeStatus"
	rsd.Open sql, conn, 0, 2, 1

	If Not rsd.EOF Then
		If rsd("OpenSlot") > 0 Then
			rsd("OpenSlot") = rsd("OpenSlot") - 1
			rsd.Update
		End If
	End If

	rsd.Close

End Function

'申請情報の取得（指定日、指定グループ）
'ユーザ名もユーザテーブルより取得
Sub GetAppInfo(conn, rsd, sGrpID, WorkDate)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	sGrpID		[ I ]グループコード
'	WorkDate	[ I ]作業日(YYYYMMDD、WAIT：引き取り待ち)

	Dim sql
	Dim sDateNow

	If WorkDate <> "WAIT" Then
		sql = _
		"SELECT sAppliInfo.*, sMUser.Name FROM sAppliInfo, sMUser" & _
		" WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'" & _
		  " AND sAppliInfo.WorkDate = '" & WorkDate & "'" & _
		  " AND RTRIM(sAppliInfo.UserID) = sMUser.UserID" & _
		" Order By sAppliInfo.DelFlag, sAppliInfo.Term, " & _
				 " sAppliInfo.OpeOrder, sAppliInfo.HopeTerm, sAppliInfo.AppDate"
	Else
		'現在日付取得
		sDateNow = GetYMDStr(Date())

		'引き取り待ち検索
		''(搬出でＳＹにあり完了または空バンでＳＹにあり完了)　あるいは
		''搬入でＳＹにあり削除かキャンセルか過去の搬入日
		sql = _
		"SELECT sAppliInfo.*, sMUser.Name as UName, sTerminal.Name as TName FROM sAppliInfo, sMUser, sTerminal" & _
		" WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'" & _
		  " AND sAppliInfo.Place = 'SY'" & _
		  " AND RTRIM(sAppliInfo.UserID) = RTRIM(sMUser.UserID)" & _
		  " AND RTRIM(sAppliInfo.Terminal) = RTRIM(sTerminal.Terminal)" & _
		  " AND ((sAppliInfo.RecDel = 'D' AND sAppliInfo.Status = '04' or sAppliInfo.RecDel = 'V' AND sAppliInfo.Status = '04' )" & _
		" OR (sAppliInfo.RecDel = 'R' AND" & " (sAppliInfo.Status = '03'" & _
		" OR sAppliInfo.DelFlag = 'Y'" & _
		" OR sAppliInfo.WorkDate < '" & sDateNow & "')))" & _
		" Order By sAppliInfo.WorkDate, sAppliInfo.Term, sAppliInfo.OpeOrder"
	End If
	rsd.Open sql, conn, 0, 1, 1

End Sub

'申請情報の取得（指定日、全グループ）
Function GetAppInfoAll(conn, rsd, WorkDate)
	'conn		[ I ]DB接続Object
	'rsd		[I/O]レコードセットObject
	'WorkDate	[ I ]作業日
	'戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
		  " Order By DelFlag, Term, OpeOrder"
	rsd.Open sql, conn, 0, 1, 1

End Function

'申請情報の取得（指定日、指定グループ）
Sub GetAppInfoDG(conn, rsd, sGrpID, WorkDate)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	sGrpID		[ I ]グループコード
'	WorkDate	[ I ]作業日

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND RTRIM(GroupID) = '" & sGrpID & "'" & _
			" AND DelFlag <> 'Y'" & _
			" AND Status = '02'" & _
		  " Order By DelFlag, Term, OpeOrder"
	rsd.Open sql, conn, 0, 1, 1

End Sub

'申請情報の取得（指定日、指定時間帯、指定グループ）
Function GetAppInfoTerm(conn, rsd, WorkDate, sHH, sGrpID)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	WorkDate	[ I ]作業日
'	sHH			[ I ]時間帯（HH）
'	sGrpID		[ I ]グループID
'	戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND Term = '" & sHH & "'" & _
		    " AND DelFlag  <> 'Y'"
	If sGrpID <> "" Then
		sql = sql & " AND RTRIM(GroupID) = '" & sGrpID & "'"
	End If
	rsd.Open sql, conn, 0, 1, 1

End Function

'申請情報の取得（指定日、指定時間帯）
Function GetAppInfoTermAll(conn, rsd, WorkDate, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	WorkDate	[ I ]作業日
'	sHH			[ I ]時間帯（HH）
'	戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND Term = '" & sHH & "'" & _
		    " AND DelFlag  <> 'Y'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'申請情報の取得（指定作業番号）
Function GetAppInfoOpeNo(conn, rsd, iOpeNo)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	iOpeNo		[ I ]作業番号
'	戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iOpeNo)
	rsd.Open sql, conn, 0, 1, 1

End Function

'申請情報の取得（指定作業番号、更新モード）
Function GetAppInfoOpeNoUpd(conn, rsd, iOpeNo)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	iOpeNo		[ I ]作業番号
'	戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iOpeNo)
	rsd.Open sql, conn, 0, 2, 1

End Function

'申請情報の取得（指定日、指定シャーシ）
Function GetAppInfoChassis(conn, rsd, WorkDate, sChassisID)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	WorkDate	[ I ]作業日
'	sChassisID	[ I ]シャーシID
'	戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND ChassisId = '" & sChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'シャトル作業数の取得（指定日、指定時間帯）
Function GetAppWorkCnt(conn, rsd, WorkDate, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	WorkDate	[ I ]作業日
'	sHH			[ I ]時間帯（HH）
'	戻り値		[ O ]シャトル作業数（往復数）

	Dim iCnt

	iCnt = 0
	'申請情報取得（指定日、指定時間帯）削除以外
	Call GetAppInfoTermAll(conn, rsd, WorkDate, sHH)
	Do Until rsd.EOF
		If rsd("DualOpeNo") = 0 Then
			iCnt = iCnt + 2
		Else
			iCnt = iCnt + 1
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	GetAppWorkCnt = iCnt / 2
End Function

'新規作業順位の取得（指定日、指定時間帯）
Function GetNewOpeOrder(conn, rsd, WorkDate, sHH, sRD)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	WorkDate	[ I ]作業日
'	sHH			[ I ]時間帯（HH）
'   sRD			[ I ]作業タイプ（R/D/V）'空バンは搬出とする
'	戻り値		[ O ]新規作業順位

	Dim iCnt

	If sHH = "12" Or sHH = "A" Or sHH = "B" Then
		GetNewOpeOrder = 0
		Exit Function
	End If

	iCnt = 0
	'申請情報取得（指定日、指定時間帯）削除以外
	Call GetAppInfoTermAll(conn, rsd, WorkDate, sHH)
	Do Until rsd.EOF
		If rsd("OpeOrder") > iCnt And _
		   rsd("OpeOrder") < 999 Then
			iCnt = rsd("OpeOrder")
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	iCnt = iCnt + 1
	If Left(sRD, 1) = "R" Then	'搬入は奇数
		If (iCnt Mod 2) = 0 Then
			iCnt = iCnt + 1
		End If
	Else						'搬出・空バンは偶数
		If (iCnt Mod 2) = 1 Then
			iCnt = iCnt + 1
		End If
	End If

	GetNewOpeOrder = iCnt
End Function

'申請の有無チェック（指定コンテナ）
''指定コンテナが有効な申請として登録中かチェックする
Function ChkAppInfoExt(conn, rsd, sContNo)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sContNo		[ I ]コンテナ番号
'	戻り値		[ O ]True:あり False:なし（削除）

	Dim sql
	Dim sDateNow

	'現在日付取得
	sDateNow = GetYMDStr(Date())

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE RTRIM(ContNo) = '" & sContNo & "'" & _
		    " AND Status = '02'" & _
		    " AND DelFlag <> 'Y'" & _
		    " AND WorkDate >= '" & sDateNow & "'"
	rsd.Open sql, conn, 0, 1, 1

	ChkAppInfoExt = not rsd.EOF
	rsd.Close
End Function

'申請の有無チェック（指定作業番号）
Function ChkAppInfoOpeNo(conn, rsd, iOpeNo)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	iOpeNo		[ I ]作業番号
'	戻り値		[ O ]True:あり False:なし

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iOpeNo)
	rsd.Open sql, conn, 0, 1, 1

	ChkAppInfoOpeNo = not rsd.EOF
	rsd.Close
End Function

'営業日の取得
Function GetBusinessDays(conn, rsd, sDateNow, iDaysCnt, sDays())
	'conn		[ I ]DB接続Object
	'rsd		[ I ]レコードセットObject
	'sDateNow	[ I ]本日
	'iDaysCnt	[ O ]営業日数
	'sDays		[ O ]営業日
	'戻り値		[ O ]なし

	Dim sql, sDate, iCnt

	'カレンダー取得
	sql = "SELECT Date FROM sCalendar WHERE HolydayFlag != 'Y' order by Date"
	rsd.Open sql, conn, 0, 1, 1

	iDaysCnt = 0
	'前営業日取得
	Do Until rsd.EOF
		sDate = Trim(rsd("Date"))
		If sDate >= sDateNow Then
			Exit Do
		End if
		sDays(iDaysCnt) = sDate
		rsd.MoveNext
	Loop
	iDaysCnt = iDaysCnt + 1

	'今日が営業日なら追加
	If sDate = sDateNow Then
		sDays(iDaysCnt) = sDate
		iDaysCnt = iDaysCnt + 1
	End if

	'６日後までの営業日取得
	iCnt = 0
	rsd.MoveFirst
	Do Until rsd.EOF
		sDate = Trim(rsd("Date"))
		If sDate > sDateNow Then
			sDays(iDaysCnt) = sDate
			iDaysCnt = iDaysCnt + 1
			iCnt = iCnt + 1
		End if
		If iCnt = 6 Then
			Exit Do
		End if
		rsd.MoveNext
	Loop
	rsd.Close

End Function

'空きスロットの取得
Function GetLackChassis(conn, rsd, sGrpID, _
			iEmptySlot, iEmptyChassis20, iEmptyChassis40)
	'conn				[ I ]DB接続Object
	'rsd				[ I ]レコードセットObject
	'sGrpID				[ I ]グループID
	'iEmptySlot			[ O ]空きスロット数
	'iEmptyChassis20	[ O ]空きシャーシ数
	'iEmptyChassis40	[ O ]空きシャーシ数

	Dim sql

	sql = "SELECT * FROM sLackChassis" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		iEmptySlot      = rsd("EmptySlot")
		iEmptyChassis20 = rsd("EmptyChassis20")
		iEmptyChassis40 = rsd("EmptyChassis40")
	Else	'空きなしの場合はレコードがない
		iEmptySlot      = 0
		iEmptyChassis20 = 0
		iEmptyChassis40 = 0
	End If
	rsd.Close

End Function

'シャトル数によるグループ割り当て枠の取得
Function GetShuttleUnit(conn, rsd, sYMD, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sYMD		[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	戻り値		[ O ]グループ割り当て枠

	Dim sql
	Dim iWk, iQty
	
	iWk = Weekday(cDate(ChgYMDStr(sYMD)))
	iQty = 0
	sql = "SELECT * FROM sShuttleUnit" & _
		  " WHERE Weekday = " & CStr(iWk)
	rsd.Open sql, conn, 0, 1, 1

	iQty = 0
	If Not rsd.EOF Then
		If sHH < "12" Then
			iQty = rsd("reservAM")
		Else
			iQty = rsd("reservPM")
		End If
	End If
	rsd.Close

	GetShuttleUnit = iQty

End Function

'グループ数の取得
Function GetGrpCnt(conn, rsd)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	戻り値		[ O ]グループ数

	Dim sql
	Dim iWk

	sql = "SELECT GroupID FROM sMGroup"
	rsd.Open sql, conn, 0, 1, 1

	iWk = 0
	Do Until rsd.EOF
		iWk = iWk + 1
		rsd.MoveNext
	Loop

	rsd.Close

	GetGrpCnt = iWk

End Function

'時間帯幅の取得
Function GetTimeSlotSize(conn, rsd, sHH, iTime)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sHH			[ I ]予約時間帯（HH）
'	iTime		[ O ]時間帯幅（分）
'	戻り値		[ O ]なし

	Dim sWk

	'時間帯幅を計算
	iTime = 60
	sWk = GetEnv(conn, rsd, "ShtStart")
	If sHH = Left(sWk, 2) Then
		iTime = 60 - CLng(Right(sWk, 2))
	End If
	sWk = GetEnv(conn, rsd, "ShtEnd")
	If sHH = Left(sWk, 2) Then
		iTime = CLng(Right(sWk, 2))
	End If
'''2001/08/20 debug
'''Response.Write "時間帯幅(" & iTime & ")"
'''
End Function

'搬出許可コンテナの取得
Function GetOutCont(conn, rsd, sContNo)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	sContNo		[ I ]コンテナ番号
'	戻り値		[ O ]なし

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sql = "SELECT * FROM sOutContainer" & sNo & _
		  " WHERE RTRIM(ContNo) = '" & sContNo & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'搬出許可マニフェストコンテナの取得
Function GetOutBLCont(conn, rsd, sBLNo)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	sBLNo		[ I ]BL番号
'	戻り値		[ O ]なし

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sql = "SELECT * FROM sOutBLCont" & sNo & _
		  " WHERE RTRIM(BLNo) = '" & sBLNo & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'搬出許可マニフェストコンテナとコンテナの取得(複数ＣＹ対応)01/10/01
Function GetOutBLContCont(conn, rsd, sBLNo, sTerm_CD)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	sBLNo		[ I ]BL番号
'	sTerm_CD	[ I ]ターミナルコード
'	戻り値		[ O ]なし

	Dim sql, sNo, sTbl1, sTbl2

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sTbl1 = "sOutContainer" & sNo
	sTbl2 = "sOutBLCont" & sNo

	sql = "SELECT " & sTbl1 & ".*" & _
		  " FROM " & sTbl1 & "," & sTbl2 & _
		  " WHERE RTRIM(" & sTbl2 & ".BLNo) = '" & sBLNo & "'" & _
		    " AND RTRIM(" & sTbl2 & ".ContNo) = RTRIM(" & sTbl1 & ".ContNo)" & _
		    " AND RTRIM(" & sTbl1 & ".Terminal) = '" & sTerm_CD & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'ブッキングの有無判定(02/02/09)
Function ChkBooking(conn, rsd, sBookNo, sOpeRec, sVslCode, sVoyage)
'	conn		[ I ]DB接続Object
'	rsd		[ I ]レコードセットObject
'	sBookNo		[ I ]ブッキング番号
'	sOpeRec		[ O ]搬入オペレータ
'	sVslCode	[ O ]搬入船名		'02/02/09 ADD
'	sVoyage		[ O ]搬入次航		'02/02/09 ADD
'	戻り値		[ O ]ブッキングの有無

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sql = "SELECT * FROM sBooking" & sNo & _
		  " WHERE RTRIM(BookNo) = '" & sBookNo & "'"
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
		sOpeRec = Trim(rsd("Operator"))
'02/02/09 ADD
		sVslCode = Trim(rsd("VslCode"))	
		sVoyage = Trim(rsd("Voyage"))
'02/02/09 ADD
	End If
	ChkBooking = Not rsd.EOF
	rsd.Close

End Function

'シャーシの取得
Function GetChassis(conn, rsd, sChassisID)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	sChassisID	[ I ]シャーシID
'	戻り値		[ O ]なし

	Dim sql

	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(ChassisId) = '" & sChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'グループの割り当て枠を取得
Function GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sGrpID		[ I ]グループコード
'	sYMD		[ I ]作業日（YYYYMMDD）
'	sHH			[ I ]時間帯（HH）
'	戻り値		[ O ]グループの割り当て枠
	Dim iCnt, iTime

	'グループ割り当て枠の取得
	iCnt = GetShuttleUnit(conn, rsd, sYMD, sHH)
'''2001/08/20 debug
'''Response.Write "グループ割り当て枠(" & iCnt & ")"
'''

	'時間帯幅の取得
	Call GetTimeSlotSize(conn, rsd, sHH, iTime)

	GetGrpAssign = iCnt * iTime / 60

'''2001/08/20 debug
'''Response.Write "割り当て枠(" & GetGrpAssign & ")"
'''

End Function

'作業番号の発番
Function GetNewOpeNo(conn, rsd, sYMD, sOpeNo)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sYMD		[ I ]シャーシID
'	sOpeNo		[ O ]作業番号
'	戻り値		[ O ]なし

	Dim bOk

	bOk = False
	Do Until bOk
		'作業番号の発番
		Call GetNewOpeNo2(conn, rsd, sYMD, sOpeNo)

		'申請の有無チェック（指定作業番号）
		If Not ChkAppInfoOpeNo(conn, rsd, CLng(sOpeNo)) Then
			bOk = True
		End If
	Loop

End Function

'作業番号の発番
Function GetNewOpeNo2(conn, rsd, sYMD, sOpeNo)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sYMD		[ I ]シャーシID
'	sOpeNo		[ O ]作業番号
'	戻り値		[ O ]なし

	Dim sql, iVal

	sql = "SELECT * FROM sMAutoNumber" & _
		  " WHERE KindCode = '01'" & _
		    " AND KeyCode = '" & Right(sYMD, 2) & "'"
	rsd.Open sql, conn, 0, 2, 1

	If rsd.EOF Then
		'新規の場合
		rsd.AddNew
		rsd("KindCode") = "01"
		rsd("KeyCode") = Right(sYMD, 2)
		rsd("InitVal") = 1
		rsd("CurrentVal") = 1
		rsd("MaxVal") = 999
		rsd("TrgDate") = sYMD
	End If

	'日付が変わった場合
	If rsd("TrgDate") <> sYMD Then
		rsd("CurrentVal") = rsd("InitVal")
		rsd("TrgDate") = sYMD
	End If

	'作業番号取得
	iVal = rsd("CurrentVal")

	If iVal = rsd("MaxVal") Then
		'最大値の場合
		rsd("CurrentVal") = rsd("InitVal")
	Else
		rsd("CurrentVal") = iVal + 1
	End If
	rsd("UpdtTime") = Now()
	rsd("UpdtPgCd") = "SYWB012"
	rsd("UpdtTmnl") = "ME"

	rsd.Update
	rsd.Close

	sOpeNo = Right(sYMD, 2) & Right("000" & CStr(iVal), 3)
	GetNewOpeNo2 = sOpeNo
End Function

'曜日を取得する
Function sWeekday(iWk)
'	sYMD		[ I ]数値
'	戻り値		[ O ]曜日

	select case	iWk
        Case 1
            sWeekday =  "日"
        Case 2
            sWeekday =  "月"
        Case 3
            sWeekday =  "火"
        Case 4
            sWeekday =  "水"
        Case 5
            sWeekday =  "木"
        Case 6
            sWeekday =  "金"
        Case 7
            sWeekday =  "土"
	end select
End Function

'予約日、予約時間帯が有効かチェック
Function CheckAppWorkDate(conn, rsd, sYMD, sHH, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]予約時間帯（HH）
'	sErrMsg			[ O ]エラーメッセージ

	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime, sWk

	sErrMsg = ""

	'*** 予約日、予約時間帯が有効かチェック ***
	'現在日付取得
	sDateNow = GetYMDStr(Date())

	If sDateNow > sYMD Then
		sErrMsg = "作業日が適切ではありません"
		Exit Function
	End If
	If sDateNow = sYMD Then	'現在日指定の場合
		'運行状況を取得
		Call GetOpeStatus(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime)
		sWk = Right("0" & CStr(iNextTime), 2)
		If sHH < sWk or (sHH = sWk and iNextStat = 2) Then
			sErrMsg = "指定時間帯の予約は締め切られています"
			Exit Function
		End If
	End If
End Function

'固定桁数での表示
Function GetNumStr(vNum, iK)
'	vNum			[ I ]数値（数値or文字列）
'	iK				[ I ]桁数
	GetNumStr = Right(String(iK, "0") & CStr(vNum), iK)
End Function

'ターミナルの取得
Function GetTerminal(conn, rsd, sTerminal())
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sTerminal	[ O ]ターミナル名称(ｺｰﾄﾞ3桁名前)
'	戻り値		[ O ]なし

	Dim sql
	Dim i

	'ターミナル取得
	sql = "SELECT * FROM sTerminal WHERE Umu <> '1' "
	sql = sql & "  Order By Terminal"
	rsd.Open sql, conn, 0, 1, 1

	i = 0
	Do Until rsd.EOF
		sTerminal(i) = rsd("Terminal") & Trim(rsd("Name"))
		i = i + 1
		rsd.MoveNext
	Loop
	rsd.Close


End Function

'コンテナタイプ名称の取得
Function GetContType(conn, rsd, sContType())
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sContType	[ O ]コンテナタイプ名称(ｺｰﾄﾞ2桁名前)
'	戻り値		[ O ]なし

	Dim sql
	Dim i

	'コンテナタイプ取得
	sql = "SELECT * FROM sContType" 
	rsd.Open sql, conn, 0, 1, 1

	i = 0
	Do Until rsd.EOF
		sContType(i) = rsd("ContType") & Trim(rsd("Name"))
		i = i + 1
		rsd.MoveNext
	Loop
	rsd.Close


End Function

'コンテナ材質取得の取得
Function GetMaterial(conn, rsd, sMaterial())
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sMaterial	[ O ]コンテナ材質(ｺｰﾄﾞ3桁名前)
'	戻り値		[ O ]なし

	Dim sql
	Dim i

	'コンテナ材質取得
	sql = "SELECT * FROM sMaterial" 
	rsd.Open sql, conn, 0, 1, 1

	i = 0
	Do Until rsd.EOF
		sMaterial(i) = rsd("Material") & Trim(rsd("Name"))
		i = i + 1
		rsd.MoveNext
	Loop
	rsd.Close

End Function


'ターミナル名称等の取得
Function GetTerminal2(conn, rsd, sTerminal, sType, sName)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sTerminal	[ I ]ターミナルコード
'	sType		[ O ]タイプ
'	sName		[ O ]ターミナル名称
'	戻り値		[ O ]なし

	Dim sql

	'ターミナル取得
	sql = "SELECT Terminal, Type, Name FROM sTerminal" & _
		  " WHERE RTRIM(Terminal) = '" & sTerminal & "'"
	rsd.Open sql, conn, 0, 1, 1

	sType  = Trim(rsd("Type"))		
	sTerm_Name  = Trim(rsd("Name"))
	rsd.Close
End Function

%>