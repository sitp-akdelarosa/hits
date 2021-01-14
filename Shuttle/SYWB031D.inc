<% 

'予約変更処理
Function UpdOpeUpd(conn, rsd, sOpeNo, sSend, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sOpeNo			[ I ]作業番号
'	sSend			[ I ]搬入元・搬出先
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
	sRDType    = Trim(rsd("RecDel"))		'搬出入区分
	rsd.Close

	If sDelFlag = "Y" Then	'削除の場合
		If sRDType = "R" and sPlace = "SY" Then
			sErrMsg = "対象は戻り作業です"
			Exit Function
		End If
		sErrMsg = ""
		Exit Function
	End If

	If sStatus = "04" Then	'完了の場合
		sErrMsg = "対象は作業を完了しています"
		Exit Function
	End If

	If sStatus = "03" Then	'完了の場合
		sErrMsg = "対象は作業をキャンセルしています"
		Exit Function
	End If

	'更新処理
	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	If sRDType = "D" Or sRDType = "V" Then	'<!--空バン対応 -->
		rsd("DeliverTo") = sSend			'搬出先
	Else
		rsd("ReceiveFrom") = sSend			'搬入元
	End If

	rsd("SendFlag") = "Y"
	rsd("UpdtTime") = Now()				'更新日時
	rsd("UpdtPgCd") = "SYWB031"			'更新ﾌﾟﾛｸﾞﾗﾑ名
	rsd("UpdtTmnl") = "WEB"				'更新端末名
	rsd.Update
	rsd.Close
	sErrMsg = ""

End Function

%>
