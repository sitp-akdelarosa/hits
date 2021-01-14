<% 

'予約削除処理
Function UpdOpeDel(conn, rsd, sOpeNo, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sOpeNo			[ I ]作業番号
'	sErrMsg			[ O ]エラーメッセージ

	Dim sStatus, sDelFlag, sLockFlag, iDualOpeNo, sRDType, sPlace

	sErrMsg = "ＤＢ更新エラー"

	'申請情報の取得（指定作業番号）
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "対象が削除されています"
		Exit Function
	End If
	sStatus    = Trim(rsd("Status"))		'状態
	sDelFlag   = Trim(rsd("DelFlag"))		'削除フラグ
	sLockFlag  = Trim(rsd("LockFlag"))		'ロックフラグ
	iDualOpeNo =      rsd("DualOpeNo")		'デュアル作業番号
	sRDType    = Trim(rsd("RecDel"))		'搬出入区分
	sPlace     = Trim(rsd("Place"))			'場所
	rsd.Close

	If sDelFlag = "Y" Then	'削除の場合
		sErrMsg = "既に削除されています"
		Exit Function
		'以下の完全削除処理は使用しない
		If sRDType = "R" and sPlace = "SY" Then
			sErrMsg = "対象は戻り作業です"
			Exit Function
		End If
		'完全に削除
		Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
		rsd.Delete
		rsd.Update
		rsd.Close
		sErrMsg = ""
		Exit Function
	End If

	If sStatus = "04" Then	'完了の場合
		sErrMsg = "対象は作業を完了しています"
		Exit Function
	End If

	If sStatus = "03" Then	'キャンセル中の場合
		sErrMsg = "対象は作業をキャンセルしています"
		Exit Function
	End If

	'以下予約中の場合（削除状態にする）
	If sLockFlag = "Y" Then	'ロック
		sErrMsg = "対象は作業が確定しています"
		Exit Function
	End If
	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	rsd("SendFlag") = "Y"
	rsd("DelFlag") = "Y"
	rsd("DualOpeNo") = 0
	rsd("OpeOrder") = 0
	rsd("UpdtTime") = Now()				'更新日時
	rsd("UpdtPgCd") = "SYWB031"			'更新ﾌﾟﾛｸﾞﾗﾑ名
	rsd("UpdtTmnl") = "WEB"				'更新端末名
	rsd.Update
	rsd.Close
	If iDualOpeNo > 0 Then	'デュアル作業番号
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
		rsd("SendFlag") = "Y"
		rsd("DualOpeNo") = 0
		rsd("UpdtTime") = Now()			'更新日時
		rsd("UpdtPgCd") = "SYWB031"		'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl") = "WEB"			'更新端末名
		rsd.Update
		rsd.Close
	End If
	sErrMsg = ""

End Function

%>
