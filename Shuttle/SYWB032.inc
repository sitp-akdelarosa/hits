<% 

'シャーシ使用予定のチェック
''指定シャーシの使用予定があるかチェックする
Function ChkAppCha(conn, rsd, sChassisID)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sChassisID	[ I ]シャーシID
'	戻り値		[ O ]True：予定なし　False：予定あり

	Dim sql
	Dim sDateNow

	'現在日付取得
	sDateNow = GetYMDStr(Date())

	sql = "SELECT OpeNo FROM sAppliInfo" & _
			" WHERE ChassisId = '" & sChassisID & "'" & _
			  " AND WorkDate >= '" & sDateNow & "'"
	rsd.Open sql, conn, 0, 1, 1
	ChkAppCha = rsd.EOF	
	rsd.close

End Function

'シャーシ属性変更によりダイヤが変わる申請を取得
Sub GetChangeApp(conn, rsd, sChassisID, iOpeNo, iDualOpeNo)
'	conn		[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sChassisID	[ I ]シャーシID
'	iOpeNo		[ O ]搬入作業番号（0：該当作業なし）
'	iDualOpeNo	[ O ]搬出作業番号（0：該当作業なし）

	Dim sql

	iOpeNo		= 0		'搬入作業番号（0：該当作業なし）
	iDualOpeNo	= 0		'搬出作業番号（0：該当作業なし）

	'指定シャーシを使用中の搬入コンテナを取得
	sql = "SELECT OpeNo,DualOpeNo FROM sAppliInfo" & _
		  " WHERE ChassisId = '" & sChassisID & "'" & _
			" AND RecDel = 'R'" & _
			" AND Status = '02'" & _
			" AND Place = 'SY'" & _
			" AND DelFlag <> 'Y'" & _
			" AND LockFlag <> 'Y'"
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then		'該当レコードあり
		iOpeNo		= rsd("OpeNo")			'搬入作業番号
		iDualOpeNo	= rsd("DualOpeNo")		'搬出作業番号
	End If
	rsd.close
End Sub

%>