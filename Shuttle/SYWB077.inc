<% 

'空きシャーシ数の取得
'	指定グループ・作業日・時間帯での作業完了時点での
'	空きシャーシ数を計算する
Function GetEmptyChassisCnt(conn, rsd, sGrpID, sYMD, sHHName, i20, i40)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sGrpID		[ I ]グループID
'	sYMD		[ I ]指定日付(YYYYMMDD)
'	sHHName		[ I ]指定時間帯（"HH"）
'	i20			[ O ]２０フィート用空シャーシ数
'	i40			[ O ]４０フィート用空シャーシ数
'	戻り値		[ O ]True:計算ＯＫ False:計算不可

	Dim sql
	Dim iUC(1)		'夕積空シャーシ数２０Ｆ，４０Ｆ（夕積にのみ使用可能）
	Dim iNC(1)		'通常空シャーシ数２０Ｆ，４０Ｆ（時間帯にも夕積にも使用可能）
	Dim sWkYmd, dtWk	'ワーク日付
	Dim iWk

	GetEmptyChassisCnt = False

	'グループの現在空きシャーシ数取得
	'初期化
	iNC(0) = 0		'通常２０ｆ（時間帯にも夕積にも使用可能）
	iNC(1) = 0		'通常４０ｆ（時間帯にも夕積にも使用可能）
	iUC(0) = 0		'夕積２０ｆ（夕積にのみ使用可能）
	iUC(1) = 0		'夕積４０ｆ（夕積にのみ使用可能）
	'シャーシ検索
	'ＳＹ蔵置中＆コンテナなし＆搬出に使える＆指定グループ
	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
		    " AND StackFlag = 'Y'" & _
		    " AND ContFlag <> 'Y'" & _
		    " AND NotDelFlag <> 'Y'"
	rsd.Open sql, conn, 0, 1, 1
	'シャーシ数集計
	Do Until rsd.EOF
		If rsd("Size20Flag") = "Y" Then
			iWk = 0		'２０Ｆ
		Else
			iWk = 1		'４０Ｆ
		End If
		If rsd("NightFlag") <> "Y" Then	'夕積でない
			iNC(iWk) = iNC(iWk) + 1	'通常（時間帯にも夕積にも使用可能）
		Else
			iUC(iWk) = iUC(iWk) + 1	'夕積（夕積にのみ使用可能）
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	sWkYmd = GetYMDStr(Date())	'本日取得
	Do While sWkYmd <= sYMD		'本日から指定日付まで計算
		'ワーク日付の申請情報を取得
		'申請情報検索ＳＱＬ作成
		sql = "SELECT * FROM sAppliInfo" & _
				" WHERE Status = '02'" & _
				  " AND DelFlag <> 'Y'" & _
				  " AND WorkFlag = ' '" & _
				  " AND NotDelFlag <> 'Y'" & _
				  " AND DualOpeNo = 0" & _
				  " AND RTRIM(GroupID) = '" & sGrpID & "'" & _
				  " AND WorkDate = '" & sWkYmd & "'"
		If sWkYmd = sYMD Then
			'指定日の場合は時間帯を条件に追加
			sql = sql & " AND Term <= '" & sHHName & "'"
		End If
		sql = sql & " ORDER BY Term, OpeOrder"

		'申請情報検索
		rsd.Open sql, conn, 0, 1, 1

		'使用台数、追加台数を計算
		Do Until rsd.EOF
			If rsd("Size20Flag") = "Y" Then
				iWk = 0		'２０Ｆ
			Else
				iWk = 1		'４０Ｆ
			End If
'''			IF rsd("RecDel") = "D" Then	'搬出（使用される）(VP対応)
			IF rsd("RecDel") = "D" OR rsd("RecDel") = "V" Then	'搬出OR空バン（使用される）
				IF rsd("Term") = "B" Then	'夕積
					'本当は夕積の場合実際にはどちらのシャーシが使用されるか分からない
					iUC(iWk) = iUC(iWk) - 1	'夕積（夕積にのみ使用可能）
				Else
					iNC(iWk) = iNC(iWk) - 1	'通常（時間帯にも夕積にも使用可能）
				End If
			Else						'搬入（空シャーシが増える）
				IF rsd("NightFlag") = "Y" Then	'夕積にのみ使用可能
					iUC(iWk) = iUC(iWk) + 1	'夕積（夕積にのみ使用可能）
				Else
					iNC(iWk) = iNC(iWk) + 1	'通常（時間帯にも夕積にも使用可能）
				End If
			End If

			rsd.MoveNext
		Loop
	
		rsd.Close

		'計算日付の加算
		dtWk = DateAdd("d", 1, CDate(ChgYMDStr(sWkYmd)))
		sWkYmd = GetYMDStr(dtWk)	'次の日取得
	Loop

	i20 = iUC(0) + iNC(0)
	i40 = iUC(1) + iNC(1)

	GetEmptyChassisCnt = True

End Function


%>