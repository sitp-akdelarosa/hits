<% 

'指定時間帯、指定グループの申請情報を取得
Sub GetAppHH(conn, rsd, _
				sGrpID, sTrgDate, TimeSlot, TimeNo, _
				sDateNow, iCurTime, iNextTime, iNextApp, _
				iRecCnt, iDelCnt, iCloseMode, _
				iAppCnt, _
				iAppOpeNo(), sAppUserNm(), sAppContNo(), _
				sAppBLNo(), sAppRecDel(), sAppStatus(), _
				sAppPlace(), sAppChassisId(), _
				sAppWorkFlag(), sAppCReason(), sAppContSize(), _
				sAppTerm(), sAppHopeTerm(), iAppOpeOrder(), _
				iAppDualOpeNo(), sAppFromTo(), sAppDelFlag(), sDelChaStock())
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sGrpID			[ I ]グループID
'	sTrgDate		[ I ]対象日("YYYYMMDD")
'	TimeSlot		[ I ]時間帯記号（例：08～16,A,B,D）
'	TimeNo			[ I ]時間帯番号（例：8～16,30,31,32）
'	sDateNow		[ I ]現在日付("YYYYMMDD")
'	iCurTime		[ I ]現在作業中時間帯（例：8～16,31）
'	iNextTime		[ I ]次作業時間帯（例：8～16,31）
'	iNextApp		[ I ]次作業申請数（５分前確定時）
'	iRecCnt			[ O ]搬入本数
'	iDelCnt			[ O ]搬出本数
'	iCloseMode		[ O ]完了モード（0：運行前　1：完了　2：運行中　3：確定　4：確定中　-1：開放中）
'	iAppCnt			[I/O]申請情報数
'	iAppOpeNo()		[ O ]作業番号
'	sAppUserNm()	[ O ]ユーザ名
'	sAppContNo()	[ O ]コンテナ番号
'	sAppBLNo()		[ O ]ＢＬ番号
'	sAppRecDel()	[ O ]搬出入区分
'	sAppStatus()	[ O ]状態
'	sAppPlace()		[ O ]場所
'   sAppChassisId()	[ O ]シャーシID
'	sAppWorkFlag()	[ O ]作業中フラグ
'	sAppCReason()	[ O ]キャンセル理由
'	sAppContSize()	[ O ]コンテナサイズ
'	sAppTerm()		[ O ]時間帯
'	sAppHopeTerm()	[ O ]希望時間帯
'	iAppOpeOrder()	[ O ]作業順位
'	iAppDualOpeNo()	[ O ]デュアル作業番号
'	sAppFromTo()	[ O ]搬出先／搬入元
'	sAppDelFlag()	[ O ]削除フラグ
'	sDelChaStock()	[ O ]搬出指定シャーシの在庫（"E"：非在庫）

 	Dim sAppFlag		'予約中の有無（Y=あり）
	Dim sNotLockFlag	'未ロックの有無（Y=あり）
	Dim	iOOrder			'作業順位
	Dim iWk				'ワーク

	'申請情報の取得（指定日、指定時間帯）
	Call GetAppHHAll(conn, rsd, sTrgDate, TimeSlot)
	
	iRecCnt = 0			'搬入本数
	iDelCnt = 0			'搬出本数
	iOOrder = 1			'作業順位
	sAppFlag = ""		'予約中の有無（Y=あり）
	sNotLockFlag = ""	'未ロックの有無（Y=あり）
	Do Until rsd.EOF
		'全搬出入本数のカウント
		If Trim(rsd("RecDel")) = "R" Then	'搬入
			iRecCnt = iRecCnt + 1		'搬入本数
		Else								'搬出
			iDelCnt = iDelCnt + 1		'搬出本数
		End If
		'予約中の有無チェック
		If Trim(rsd("Status")) = "02" Then	'状態
			sAppFlag = "Y"			'予約中の有無（Y=あり）
		End If
		'未ロックの有無チェック
		If Trim(rsd("LockFlag")) <> "Y" Then	'状態
			sNotLockFlag = "Y"		'未ロックの有無（Y=あり）
		End If
		'作業順位の計算
		iWk = rsd("OpeOrder")		'作業順位
		If iWk > 0 And iWk < 999 Then	'作業順位が有効
			''時間帯予約でない OR キャンセル を除外
			iWk = iOOrder
			iOOrder = iOOrder + 1	'作業順位
		End If
		'対象グループの場合はレコードを記憶
		If Trim(rsd("GroupID")) = sGrpID Then
			iAppOpeNo(iAppCnt)      = rsd("OpeNo")				'作業番号
			sAppUserNm(iAppCnt)     = Trim(rsd("Name"))			'ユーザ名
			sAppContNo(iAppCnt)		= Trim(rsd("ContNo"))		'コンテナ番号
			sAppBLNo(iAppCnt)		= Trim(rsd("BLNo"))			'ＢＬ番号
			sAppRecDel(iAppCnt)     = Trim(rsd("RecDel"))		'搬出入区分
			sAppStatus(iAppCnt)     = Trim(rsd("Status"))		'状態
			sAppPlace(iAppCnt)      = Trim(rsd("Place"))		'場所
			sAppChassisId(iAppCnt)  = Trim(rsd("ChassisId"))	'シャーシID
			sAppWorkFlag(iAppCnt)   = Trim(rsd("WorkFlag"))		'作業中フラグ
			sAppCReason(iAppCnt)    = Trim(rsd("CReason"))		'キャンセル理由
			iAppDualOpeNo(iAppCnt)  = rsd("DualOpeNo")			'デュアル作業番号
			sAppHopeTerm(iAppCnt)  	= Trim(rsd("HopeTerm"))		'希望時間帯
			sAppContSize(iAppCnt)  	= Trim(rsd("ContSize"))		'コンテナサイズ
			sAppDelFlag(iAppCnt)	= Trim(rsd("DelFlag"))		'削除フラグ
			If sAppRecDel(iAppCnt) = "D" Then	'搬出時
				sAppFromTo(iAppCnt) = Trim(rsd("DeliverTo"))	'搬出先
			Else								'搬入時
				sAppFromTo(iAppCnt) = Trim(rsd("ReceiveFrom"))	'搬入元
			End If
			sDelChaStock(iAppCnt) = ""							'搬出指定シャーシの在庫

			sAppTerm(iAppCnt) 		= TimeSlot					'時間帯
			iAppOpeOrder(iAppCnt)	= iWk						'作業順位

			iAppCnt = iAppCnt + 1		'申請情報数
		End If
		rsd.MoveNext
	Loop
	rsd.Close


	'搬出予約でシャーシ指定の場合はシャーシのＳＹ在庫をチェック
	If sTrgDate <> "WAIT" Then
		For iWk = 0 To iAppCnt - 1
			If sAppRecDel(iWk) = "D" and _
			   sAppDelFlag(iWk) <> "Y" and _
			   sAppWorkFlag(iWk) <> "Y" and _
		       iAppDualOpeNo(iWk) = 0 and _
			   sAppStatus(iWk) = "02" and _
		       sAppChassisId(iWk) <> "" Then
				'空シャーシＳＹ在庫の判定
				If Not ChkEmpChaStock(conn, rsd, sAppChassisId(iWk)) Then	'なし
					sDelChaStock(iWk) = "E"	'搬出指定シャーシの在庫＝非在庫
				End If
			End If
		Next
	End If

	'完了モードの計算（0：運行前　1：完了　2：運行中　3：確定　4：確定中　-1：開放中）
	iCloseMode = 0	'運行前
	If sTrgDate = sDateNow Then			'表示が本日の場合
		If TimeNo = 12 Then				'午前予約の場合
			If 11 < iNextTime Then		'次が11時以降の場合
				iCloseMode = 1	'完了
			ElseIf 11 = iNextTime Then	'次が11時の場合
				If iNextStat = 2 Then	'５分前
					iCloseMode = 1	'完了
				End If
			End If
		ElseIf TimeNo = 30 Then			'午後予約の場合
			If iNextTime = 31 Then			'次が夕積の場合
				iCloseMode = 1	'完了
			ElseIf iNextTime = 15 Then		'15時の場合
				If iNextStat = 2 Then		'５分前
					iCloseMode = 1	'完了
				End If
			End If
		ElseIf TimeNo < iNextTime Then	'運行時間以前の場合
			If sAppFlag = "Y" Then		'予約中の有無（Y=あり）
				iCloseMode = 2	'運行中
			Else
				iCloseMode = 1	'完了
			End If
		ElseIf TimeNo = iNextTime Then	'次運行の場合
			If iNextStat = 1 Then		'３０分前
				iCloseMode = -1	'開放中
			ElseIf iNextStat = 2 Then	'５分前
				iCloseMode = 3	'確定
				'未ロック有りか確定本数が異なる場合
				If sNotLockFlag = "Y" Or _
				   (iRecCnt + iDelCnt) <> iNextApp Then
					iCloseMode = 4	'確定中
				End If
			End If
		End If
	ElseIf sTrgDate < sDateNow Then		'過去の場合
		iCloseMode = 1	'完了
	End If

End Sub

'申請情報の取得（指定日、指定時間帯）
'　ユーザ名もユーザテーブルより取得
Sub GetAppHHAll(conn, rsd, WorkDate, Term)
'conn		[ I ]DB接続Object
'rsd		[I/O]レコードセットObject
'WorkDate	[ I ]作業日
'Term		[ I ]時間帯（HH）
'戻り値		[ O ]なし

	Dim 	sql

	If Term = "D" Then	'ユーザ削除の取得
		'指定日の削除レコード取得
		sql = "SELECT sAppliInfo.*, sMUser.Name FROM sAppliInfo, sMUser" & _
			  " WHERE sAppliInfo.WorkDate = '" & WorkDate & "'" & _
			    " AND sAppliInfo.DelFlag = 'Y'" & _
				" AND RTRIM(sAppliInfo.UserID) = sMUser.UserID"
	Else				'通常の取得
		'指定日、指定時間帯の削除でないレコード取得
		sql = "SELECT sAppliInfo.*, sMUser.Name FROM sAppliInfo, sMUser" & _
			  " WHERE sAppliInfo.WorkDate = '" & WorkDate & "'" & _
			    " AND sAppliInfo.Term = '" & Term & "'" & _
			    " AND sAppliInfo.DelFlag <> 'Y'" & _
				" AND RTRIM(sAppliInfo.UserID) = sMUser.UserID"
	End If
	sql = sql & " Order By sAppliInfo.DelFlag, sAppliInfo.Term, " & _
				" sAppliInfo.OpeOrder, sAppliInfo.HopeTerm, sAppliInfo.AppDate"
	rsd.Open sql, conn, 0, 1, 1

End Sub

'指定グループの引き取り待ち申請を取得
Sub GetAppWait(conn, rsd, _
				sGrpID, _
				iAppCnt, _
				iAppOpeNo, sAppUserNm, sAppContNo, _
				sAppBLNo, sAppRecDel, sAppStatus, _
				sAppPlace, sAppChassisId, _
				sAppWorkFlag, sAppCReason, sAppContSize, _
				sAppTerm, sAppHopeTerm, iAppOpeOrder, _
				iAppDualOpeNo, sAppFromTo, sAppDelFlag)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sGrpID			[ I ]グループID
'	iAppCnt			[ O ]申請情報数
'	iAppOpeNo()		[ O ]作業番号
'	sAppUserNm()	[ O ]ユーザ名
'	sAppContNo()	[ O ]コンテナ番号
'	sAppBLNo()		[ O ]ＢＬ番号
'	sAppRecDel()	[ O ]搬出入区分
'	sAppStatus()	[ O ]状態
'	sAppPlace()		[ O ]場所
'   sAppChassisId()	[ O ]シャーシID
'	sAppWorkFlag()	[ O ]作業中フラグ
'	sAppCReason()	[ O ]キャンセル理由
'	sAppContSize()	[ O ]コンテナサイズ
'	sAppTerm()		[ O ]時間帯
'	sAppHopeTerm()	[ O ]希望時間帯
'	iAppOpeOrder()	[ O ]作業順位
'	iAppDualOpeNo()	[ O ]デュアル作業番号
'	sAppFromTo()	[ O ]搬出先／搬入元
'	sAppDelFlag()	[ O ]削除フラグ

	iAppCnt = 0		'申請情報数

	'申請情報の取得（指定日、指定グループ）
	Call GetAppInfo(conn, rsd, sGrpID, "WAIT")
	Do Until rsd.EOF
		iAppOpeNo(iAppCnt)      = rsd("OpeNo")				'作業番号
		sAppUserNm(iAppCnt)     = Trim(rsd("Name"))			'ユーザ名
		sAppContNo(iAppCnt)		= Trim(rsd("ContNo"))		'コンテナ番号
		sAppBLNo(iAppCnt)		= Trim(rsd("BLNo"))			'ＢＬ番号
		sAppRecDel(iAppCnt)     = Trim(rsd("RecDel"))		'搬出入区分
		sAppStatus(iAppCnt)     = Trim(rsd("Status"))		'状態
		sAppPlace(iAppCnt)      = Trim(rsd("Place"))		'場所
		sAppChassisId(iAppCnt)  = Trim(rsd("ChassisId"))	'シャーシID
		sAppWorkFlag(iAppCnt)   = Trim(rsd("WorkFlag"))		'作業中フラグ
		sAppCReason(iAppCnt)    = Trim(rsd("CReason"))		'キャンセル理由
		iAppDualOpeNo(iAppCnt)  = rsd("DualOpeNo")			'デュアル作業番号
		sAppHopeTerm(iAppCnt)  	= Trim(rsd("HopeTerm"))		'希望時間帯
		sAppContSize(iAppCnt)  	= Trim(rsd("ContSize"))		'コンテナサイズ
		sAppDelFlag(iAppCnt)	= Trim(rsd("DelFlag"))		'削除フラグ
		If sAppRecDel(iAppCnt) = "D" Then	'搬出時
			sAppFromTo(iAppCnt) = Trim(rsd("DeliverTo"))	'搬出先
		Else								'搬入時
			sAppFromTo(iAppCnt) = Trim(rsd("ReceiveFrom"))	'搬入元
		End If

		sAppTerm(iAppCnt) 		= Trim(rsd("Term"))			'時間帯
		iAppOpeOrder(iAppCnt)	= 0							'作業順位

		iAppCnt = iAppCnt + 1		'申請情報数

		rsd.MoveNext
	Loop
	rsd.Close

End Sub

'シャーシ設定
Sub SetAppChas( iAppCnt, _
				iAppOpeNo(), sAppUserNm(), sAppContNo(), _
				sAppBLNo(), sAppRecDel(), sAppStatus(), _
				sAppPlace(), sAppChassisId(), _
				sAppWorkFlag(), sAppCReason(), sAppContSize(), _
				sAppTerm(), sAppHopeTerm(), iAppOpeOrder(), _
				iAppDualOpeNo(), sAppFromTo())
'	iAppCnt			[ I ]申請情報数
'	iAppOpeNo()		[ I ]作業番号
'	sAppUserNm()	[ I ]ユーザ名
'	sAppContNo()	[ I ]コンテナ番号
'	sAppBLNo()		[ I ]ＢＬ番号
'	sAppRecDel()	[ I ]搬出入区分
'	sAppStatus()	[ I ]状態
'	sAppPlace()		[ I ]場所
'   sAppChassisId()	[I/O]シャーシID
'	sAppWorkFlag()	[ I ]作業中フラグ
'	sAppCReason()	[ I ]キャンセル理由
'	sAppContSize()	[ I ]コンテナサイズ
'	sAppTerm()		[ I ]時間帯
'	sAppHopeTerm()	[ I ]希望時間帯
'	iAppOpeOrder()	[ I ]作業順位
'	iAppDualOpeNo()	[ I ]デュアル作業番号
'	sAppFromTo()	[ I ]搬出先／搬入元

	Dim i, k

	'デュアルで搬入側シャーシが決定している場合に搬出側にシャーシをセット
	For i = 0 To iAppCnt - 1
		If sAppRecDel(i) = "R" And _
		   sAppChassisId(i) <> "" And _
		   iAppDualOpeNo(i) > 0 And _
		   sAppStatus(i) = "02" Then
			'デュアルの搬出側を検索
			For k = 0 To iAppCnt - 1
				If iAppOpeNo(k) = iAppDualOpeNo(i) Then
					'搬出側シャーシにセット
					sAppChassisId(k) = sAppChassisId(i)
					Exit For
				End If
			Next
		End If
	Next

End Sub

'時間帯セルの設定
Sub SetCell01(conn, rsd, sTrgDate, TimeSlot, _
				sSTime, sETime, iSTime, iETime, _
				iCloseMode, sOpenFlag(), _
				TimeName, TimeJmp, sStatus)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sTrgDate	[ I ]指定日("YYYYMMDD")
'	TimeSlot	[ I ]時間帯記号（例：08～16,A,B,D）
'	sSTime		[ I ]開始時間（HHMM）
'	sETime		[ I ]終了時間（HHMM）
'	iSTime		[ I ]開始時間帯
'	iETime		[ I ]終了時間帯
'	iCloseMode	[ I ]完了モード（0：運行前　1：完了　2：運行中　3：確定　4：確定中　-1：開放中）
'	sOpenFlag	[ I ]開放フラグ
'	TimeName	[ O ]時間帯名
'	TimeJmp		[ O ]時間帯ジャンプ
'	sStatus		[ O ]時間帯状態

	Dim iTimeNo		'時間帯番号（0～23）

	'時間帯名取得
	iTimeNo = -1
	If TimeSlot = "12" Then		'午前指定
		TimeName = "午前指定"
		If iCloseMode = 0 Then	'運行前
			sStatus = "受付中"
		Else					'運行中、運行後
			sStatus = "締切"
		End If
	ElseIf TimeSlot = "A" Then	'午後指定
		TimeName = "午後指定"
		If iCloseMode = 0 Then	'運行前
			sStatus = "受付中"
		Else					'運行中、運行後
			sStatus = "締切"
		End If
	ElseIf TimeSlot = "B" Then	'夕積指定
		TimeName = "夕積指定"
		sStatus = "受付中"
		If iCloseMode = 1 Then			'完了
			sStatus = "完了"
		ElseIf iCloseMode = 2 Then		'運行中
			sStatus = "運行中"
		ElseIf iCloseMode = 3 Then		'確定
			sStatus = "確定"
		ElseIf iCloseMode = 4 Then		'確定中
			sStatus = "確定中"
		End If
	ElseIf TimeSlot = "D" Then	'ユーザ削除
		TimeName = "ユーザ削除"
		sStatus = "　"
	Else						'時間帯指定
		'表示文字列計算
		iTimeNo = CLng(TimeSlot)
		If iTimeNo < 12 Then	'午前
			If iTimeNo = iSTime Then
				TimeName = GetTimeSlot(CLng(Left(sSTime, 2)), _
									   CLng(Right(sSTime, 2)), "S")
			Else
				TimeName = GetTimeSlot(iTimeNo, 0, "S")
			End If
		Else					'午後
			If iTimeNo = iETime Then
				TimeName = GetTimeSlot(CLng(Left(sETime, 2)), _
									   CLng(Right(sETime, 2)), "E")
			Else
				TimeName = GetTimeSlot(iTimeNo, 0, "S")
			End If
		End If

		'状態計算
		'グループ枠開放画面へのジャンプを設定
		sStatus = "<A href=""SYWB019.asp?YMD=" & sTrgDate & TimeSlot & "&NAME=" & TimeName & " "">"
		If sOpenFlag(iTimeNo) = "Y" Then
			sStatus = sStatus & "開放予定"
		Else
			sStatus = sStatus & "開放しない"
		End If

		Select Case iCloseMode
   			Case 1					'完了
				sStatus = "完了"
    		Case 2					'運行中
				sStatus = "運行中"
    		Case -1					'開放中
				If sOpenFlag(iTimeNo) = "Y" Then
					sStatus = "開放中"
				Else
					sStatus = "開放しない"
				End If
    		Case 3					'確定
				sStatus = "確定"
    		Case 4					'確定中
				sStatus = "確定中"
		End Select 
	End If

	'予約へのジャンプ
	TimeJmp = ""
	if iCloseMode < 1 and TimeSlot <> "D" Then		'未完了の場合
		'時間帯から予約へのジャンプを設定
		TimeJmp = "<A href=""SYWB010.asp?YMD=" & sTrgDate & TimeSlot & _
												"&NAME=" & TimeName & " "">"
	End If

End Sub

'データセルの設定
Sub SetCell05(iRecIdx, iCloseMode, _
				sTrgDate, TimeSlot, TimeName, _
				iAppOpeNo(),  sAppUserNm(), sAppContNo(), sAppBLNo(), _
          		sAppRecDel(), sAppStatus(),  sAppPlace(), _
          		sAppChassisId(), sAppWorkFlag(), sAppCReason(), _
          		sAppContSize(), sAppTerm(), sAppHopeTerm(), _
				iAppOpeOrder(), iAppDualOpeNo(), sAppFromTo(), _
				sAppDelFlag(), sDelChaStock(), _
				sCell())

'	iRecIdx			[ I ]レコードインデックス
'	iCloseMode		[ I ]完了モード
'	sTrgDate		[ I ]日付
'	TimeSlot		[ I ]時間帯
'	TimeName		[ I ]時間帯名
'	iAppOpeNo		[ I ]作業番号
'	sAppUserNm		[ I ]ユーザ名
'	sAppContNo		[ I ]コンテナ番号
'	sAppBLNo		[ I ]ＢＬ番号
'	sAppRecDel		[ I ]搬出入区分
'	sAppStatus		[ I ]状態
'	sAppPlace		[ I ]場所
'	sAppChassisId	[ I ]シャーシID
'	sAppWorkFlag	[ I ]作業中フラグ
'	sAppCReason		[ I ]自動削除理由
'	sAppContSize	[ I ]コンテナサイズ
'	sAppTerm		[ I ]決定時間帯（HH）
'	sAppHopeTerm	[ I ]予約時時間帯（HH）
'	iAppOpeOrder	[ I ]作業順位
'	iAppDualOpeNo	[ I ]デュアル作業番号
'	sAppFromTo		[ I ]搬出先／搬入元
'	sAppDelFlag		[ I ]削除フラグ
'	sDelChaStock	[ I ]搬出指定シャーシの在庫（"E"：在庫なし）
'	sCell			[ O ]セル内容
	Dim sOpeNo		'作業番号

	If iRecIdx = -1 Then	'対応レコードなし＝空白行
		sCell(0) = "　"
		sCell(1) = "　"
		sCell(2) = "　"
		sCell(3) = "　"
		sCell(4) = "　"
		sCell(5) = "　"
		sCell(6) = "　"
		sCell(7) = "　"
		sCell(8) = "　"
		sCell(9) = "　"
		sCell(10) = "　"
		Exit Sub
	End If
	If iRecIdx < 0 Then		'単独搬出入表示用の空白行
		sCell(0) = "－"
		sCell(1) = "－"
		sCell(2) = "－"
		sCell(3) = "－"
		sCell(4) = "－"
		sCell(5) = "－"
		sCell(6) = "－"
		sCell(7) = "－"
		sCell(8) = "－"
		sCell(9) = "－"
		sCell(10) = "－"
		If iRecIdx = -2 Then
			sCell(5) = "入"
		Else
			sCell(5) = "出"
		End If
		Exit Sub
	End If

	'表示レコードあり
		'ユーザ名
		sCell(0) = sAppUserNm(iRecIdx)

		'作業番号
		sOpeNo = GetNumStr(iAppOpeNo(iRecIdx), 5)	'固定桁数での表示
		sCell(1) = sOpeNo
		If iAppDualOpeNo(iRecIdx) > 0 Then
			sCell(1) = sCell(1) & "*"
		End If

		'作業順位
		sCell(2) = CStr(iAppOpeOrder(iRecIdx))
		If sCell(2) = "0" Or _
		   sCell(2) = "999" Then
			sCell(2) = "-"
		End If

		'コンテナ番号
		sCell(3) = ""
		If sTrgDate <> "WAIT" Then
			If TimeSlot > "B" Or _
			   sAppStatus(iRecIdx) = "03" Or _
			   iCloseMode < 1 Then
				sCell(3) = "<A href=""SYWB020.asp?YMD=" & sTrgDate & TimeSlot & _
							"&NAME=" & TimeName & _
							"&OPENO=" & sOpeNo & " "">"
			End If
		End If
		If sAppContNo(iRecIdx) <> "" Then
			sCell(3) = sCell(3) & sAppContNo(iRecIdx)
		Else
			sCell(3) = sCell(3) & sAppBLNo(iRecIdx)
		End If
		sCell(3) = sCell(3) & "</A>"

		'予約タイプ
		Select Case sAppHopeTerm(iRecIdx) 
    		Case "12"
				sCell(4) = "午前"
    		Case "A"
				sCell(4) = "午後"
    		Case "B"
				sCell(4) = "夕積"
    		Case Else 
				sCell(4) = "時間"
		End Select 

		'入／出区分
		If sAppRecDel(iRecIdx) = "R" Then
			sCell(5) = "入"
		Else
			sCell(5) = "出"
		End If

		'コンテナサイズ
		sCell(6) = sAppContSize(iRecIdx)

		'シャーシID
		If sTrgDate <> "WAIT" Then
			'搬入の場合はシャーシ属性変更呼出が可能
			If sAppChassisId(iRecIdx) = "" Then
				sCell(7) = "-"
			Else
				'シャーシがある場合のみ
				sCell(7) = "<A href=""SYWB024.asp?YMD=" & sTrgDate & _
												"&TRGID=" & sAppChassisId(iRecIdx) & _
												" "">"
			End If

			If sAppRecDel(iRecIdx) = "D" Then
				'搬出の場合は使用シャーシ変更呼出が可能
				If iCloseMode = 0 And _
					TimeSlot <> "D" Then
					sCell(7) = "<A href=""SYWB022.asp?YMD=" & sTrgDate & _
												"&OPENO=" & sOpeNo & _
												"&CID=" & sAppChassisId(iRecIdx) & _
												" "">"
					If sAppChassisId(iRecIdx) = "" Then
						sCell(7) = sCell(7) & "-"
					End If
				End If
			End If

			sCell(7) = sCell(7) & sAppChassisId(iRecIdx) & "</A>"
		Else
			If sAppChassisId(iRecIdx) = "" Then
				sCell(7) = "　"
			Else
				sCell(7) = sAppChassisId(iRecIdx)
			End If
		End If

		'場所
		If sAppPlace(iRecIdx) = "" Then
			sCell(8) = "-"
		ElseIf sAppPlace(iRecIdx) = "MV" Then
			sCell(8) = "-"
		Else
			sCell(8) = sAppPlace(iRecIdx)
		End If

		'状態、備考
		sCell(9) = "　"
		sCell(10) = "　"

		'搬出予約で指定シャーシがＳＹ在庫でない場合
		If sDelChaStock(iRecIdx) = "E" Then
			sCell(10) = "ｼｬｰｼ無"
		End If

		If sAppDelFlag(iRecIdx) = "Y" Then		'ユーザ削除
			'削除中
			sCell(9) = "削除"
			If sAppPlace(iRecIdx) = "SY" Then
				sCell(10) = "引取待"
			End If
		Else
			Select Case sAppStatus(iRecIdx) 
    			Case "02"	'未完了
					If sAppWorkFlag(iRecIdx) = "Y" Then
						sCell(9) = "作業中"
					Else
						If sAppFromTo(iRecIdx) = "" Then
							If sAppRecDel(iRecIdx) = "R" Then
								sCell(10) = "搬入元無"
							Else
								sCell(10) = "搬出先無"
							End If
						End If
						If sAppRecDel(iRecIdx) = "R" and _
						   sAppPlace(iRecIdx) <> "SY" Then
							sCell(10) = "未着"
						End If
						sCell(9) = "待"
						Select Case sAppTerm(iRecIdx) 
    						Case "12", "A"
								sCell(9) = "未定"
    						Case "B"
								'未定でもあり待ちでもある
						End Select 
					End If
    			Case "03"	'キャンセル
					sCell(9) = "削除"
					Select Case sAppCReason(iRecIdx) 
    					Case "1"	'未着
							sCell(10) = "未着"
    					Case "2"	'BLなし
							sCell(10) = "BLなし"
    					Case "3"	'シャーシ不足
							sCell(10) = "ｼｬｰｼ不足"
    					Case "4"	'指定シャーシがない
							sCell(10) = "ｼｬｰｼ無"
    					Case "5"	'ゲート拒否
							If sAppRecDel(iRecIdx) = "R" Then
								sCell(10) = "搬入不可"
							Else
								sCell(10) = "搬出失敗"
							End If
    					Case "6"	'非在庫
							sCell(10) = "非在庫"
    					Case "7"	'時間不足
							sCell(10) = "能力超"
    					Case "8"	'時間不足
							If sAppRecDel(iRecIdx) = "R" Then
								sCell(10) = "搬入元無"
							Else
								sCell(10) = "搬出先無"
							End If
    					Case Else 
							sCell(10) = "不明"
					End Select 
					If sAppRecDel(iRecIdx) = "D" and _
					   sAppPlace(iRecIdx) = "SY" Then
						sCell(10) = "戻し"
					End If
    			Case "04"	'完了
					sCell(9) = "完了"
					If sAppRecDel(iRecIdx) = "D" and _
					   sAppPlace(iRecIdx) = "SY" Then
						sCell(10) = "引取待"
					End If
    			Case Else 
					sCell(9) = "不明"
			End Select 
		End If

End Sub

'作業中申請の有無チェック
Function ChkWorkingApp(conn, rsd, WorkDate, sHH)
'	conn		[ I ]DB接続Object
'	rsd			[I/O]レコードセットObject
'	WorkDate	[ I ]作業日
'	sHH			[ I ]時間帯（HH）
'	戻り値		[ O ]True：あり　False：なし

	ChkWorkingApp = False

	'申請情報の取得（指定日、指定時間帯、指定グループ）
	Call GetAppInfoTerm(conn, rsd, WorkDate, sHH, "")
	Do Until rsd.EOF
		If Trim(rsd("Status")) = "02" Or _
		   Trim(rsd("WorkFlag")) = "Y" Then
			ChkWorkingApp = True			
		End If
		rsd.MoveNext
	Loop
	rsd.Close
End Function

'データセルカラーの計算
Sub	CalcDataColor(sColor, sCell())
'	sColor		[ O ]データセルカラー
'	sCell()		[ I ]セルの値

	sColor = ""
	If sCell(10) = "引取待" Then
		sColor = "bgcolor=""#f0f0ff"" "
	End If
	If sCell(10)= "未着" or _
	   sCell(10)= "搬入元無" or _
	   sCell(10)= "搬出先無" or _
	   sCell(10)= "ｼｬｰｼ無" then
		sColor = "bgcolor=""#FFFF00"" "
	End If
	If sCell(9) = "削除" Then
		sColor = "bgcolor=""#dda0dd"" "
	End If
	If sCell(9) = "作業中" Then
		sColor = "bgcolor=""#F08080"" "
	End If
End Sub

'仮シャーシ計算
''３０分前に単独搬出のシャーシを仮に表示する
Sub CalcAppChas(conn, rsd, _
				sGrpID, sTrgDate, _
				iCurTime, iNextTime, iNextStat, _
				iAppCnt, _
				iAppOpeNo(), sAppUserNm(), sAppContNo(), _
				sAppBLNo(), sAppRecDel(), sAppStatus(), _
				sAppPlace(), sAppChassisId(), _
				sAppWorkFlag(), sAppCReason(), sAppContSize(), _
				sAppTerm(), sAppHopeTerm(), iAppOpeOrder(), _
				iAppDualOpeNo(), sAppFromTo())
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sGrpID			[ I ]グループID
'	sTrgDate		[ I ]対象日("YYYYMMDD")
'	iCurTime		[ I ]現作業時間帯（例：8～16,31）
'	iNextTime		[ I ]次作業時間帯（例：8～16,31）
'	iNextStat		[ I ]次作業状況（1：３０分前　2：５分前）
'	iAppCnt			[ I ]申請情報数
'	iAppOpeNo()		[ I ]作業番号
'	sAppUserNm()	[ I ]ユーザ名
'	sAppContNo()	[ I ]コンテナ番号
'	sAppBLNo()		[ I ]ＢＬ番号
'	sAppRecDel()	[ I ]搬出入区分
'	sAppStatus()	[ I ]状態
'	sAppPlace()		[ I ]場所
'   sAppChassisId()	[I/O]シャーシID
'	sAppWorkFlag()	[ I ]作業中フラグ
'	sAppCReason()	[ I ]キャンセル理由
'	sAppContSize()	[ I ]コンテナサイズ
'	sAppTerm()		[ I ]時間帯
'	sAppHopeTerm()	[ I ]希望時間帯
'	iAppOpeOrder()	[ I ]作業順位
'	iAppDualOpeNo()	[ I ]デュアル作業番号
'	sAppFromTo()	[ I ]搬出先／搬入元

	Dim iCCnt				'シャーシ数
	Dim sChassisId(300)		'シャーシID
	Dim sSize20Flag(300)	'20フィートフラグ
	Dim sMixSizeFlag(300)	'20/40兼用シャーシ
	Dim sNightFlag(300)		'夕積フラグ
	Dim i, k
	Dim sCurTime			'現作業時間帯
	Dim sNextTime			'次作業時間帯
	Dim sWk

	'時間帯番号から時間帯記号を取得
	sCurTime  = TermNumToSym(iCurTime)
	If iNextStat > 0 Then	'３０分前以降の場合のみ対象とする
		sNextTime = TermNumToSym(iNextTime)
	Else
		sNextTime = ""
	End If

	'使用可能空シャーシの取得
	Call GetEmpCha(conn, rsd, sGrpID, sTrgDate, _
				iCCnt, sChassisId, sSize20Flag, sMixSizeFlag, sNightFlag)

	For i = 0 To iAppCnt - 1
		'単独搬出でシャーシ指定がない申請に対して計算
		If sAppRecDel(i) = "D" And _
		   sAppChassisId(i) = "" And _
		   iAppDualOpeNo(i) = 0 And _
		   sAppStatus(i) = "02" And _
		   (sAppTerm(i) = sCurTime Or sAppTerm(i) = sNextTime) Then
			If sAppContSize(i) = "20" Then	'申請が２０フィートの場合
				sWk = "Y"
			Else							'申請が４０フィートの場合
				sWk = ""
			End If
			'シャーシを検索
			For k = 0 To iCCnt - 1
				If sChassisId(k) <> "" Then	'有効なシャーシ
					'サイズが適合するか兼用シャーシ
					If sSize20Flag(k) = sWk Or sMixSizeFlag(k) = "Y" Then
						'夕積予約のみ夕積シャーシを使用可能
						If sAppTerm(i) = "B" Or sNightFlag(k) <> "Y" Then
							sAppChassisId(i) = sChassisId(k)
							sChassisId(k) = ""
							Exit For
						End If
					End If
				End If
			Next
		End If
	Next
End Sub

'使用可能空シャーシの取得
Sub GetEmpCha(conn, rsd, sGrpID, sTrgDate, _
				iCCnt, sChassisId(), _
				sSize20Flag(), sMixSizeFlag(), sNightFlag())
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sGrpID			[ I ]グループID
'	sTrgDate		[ I ]対象日("YYYYMMDD")
'	iCCnt			[ O ]シャーシ数
'	sChassisId()	[ O ]シャーシID
'	sSize20Flag()	[ O ]20フィートフラグ
'	sMixSizeFlag()	[ O ]20/40兼用シャーシ
'	sNightFlag()	[ O ]夕積フラグ

	Dim sql
	Dim i, sWk

	'SELECT文作成
	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
		    " AND StackFlag = 'Y'" & _
		    " AND ContFlag <> 'Y'" & _
		    " AND NotDelFlag <> 'Y'" & _
		  " Order By ChassisId"

	'検索
	rsd.Open sql, conn, 0, 1, 1

	'レコード記憶
	iCCnt = 0
	Do Until rsd.EOF
		sChassisId(iCCnt)   = Trim(rsd("ChassisId"))	'シャーシID
		sSize20Flag(iCCnt)  = Trim(rsd("Size20Flag"))	'20フィートフラグ
		sMixSizeFlag(iCCnt) = Trim(rsd("MixSizeFlag"))	'20/40兼用シャーシ
		sNightFlag(iCCnt)   = Trim(rsd("NightFlag"))	'夕積フラグ
		iCCnt = iCCnt + 1	
		rsd.MoveNext
	Loop
	rsd.Close

	'申請情報の取得（指定日、指定グループ）
	Call GetAppInfoDG(conn, rsd, sGrpID, sTrgDate)
	'本日の申請にあるシャーシは候補から除外する
	Do Until rsd.EOF
		sWk = Trim(rsd("ChassisId"))			'シャーシID
		For i = 0 To iCCnt - 1
			If sChassisId(i) = sWk Then
				sChassisId(i) = ""
			End If
		Next
		rsd.MoveNext
	Loop
	rsd.Close

End Sub

'空シャーシＳＹ在庫の判定
Function ChkEmpChaStock(conn, rsd, sChassisId)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sChassisId		[ I ]シャーシID
'	ReturnCode		[ O ]True：シャーシあり　False：シャーシなし
	Dim sql

	'SELECT文作成
	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(ChassisId) = '" & sChassisId & "'" & _
		    " AND StackFlag = 'Y'" & _
		    " AND ContFlag <> 'Y'"

	'検索
	rsd.Open sql, conn, 0, 1, 1

	ChkEmpChaStock = Not rsd.EOF
	rsd.Close

End Function

%>