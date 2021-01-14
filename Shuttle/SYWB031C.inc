<% 

'予約交換処理
Function UpdOpeChg(conn, rsd, sOpeNo, sYMD, sChgOpeNo, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sOpeNo			[ I ]作業番号
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sChgOpeNo		[ I ]交換対象作業番号
'	sErrMsg			[ O ]エラーメッセージ

	Dim sGrpID(1), sUsrID(1), dWorkDate(1), sHH(1)
	Dim iOpeOrder(1), sStatus(1), sDelFlag(1), sLockFlag(1), iDualOpeNo(1)
	Dim sRDType(1), sContSize(1), bChkA(1), bChkB(1), bChkC(1), sChID(1), sPlace(1)
	Dim AppDate(1), sTerminal(1)

	sErrMsg = "ＤＢ更新エラー"

	'申請情報の取得（指定作業番号）
	Call GetAppInfoOpeNo(conn, rsd, CLng(sChgOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "交換相手が存在しません"
		Exit Function
	End If
	dWorkDate(0)  =      rsd("WorkDate")	'作業日
	sHH(0)    	  = Trim(rsd("Term"))		'時間帯
	iOpeOrder(0)  =      rsd("OpeOrder")	'作業順位
	sStatus(0)    = Trim(rsd("Status"))		'状態
	sDelFlag(0)   = Trim(rsd("DelFlag"))	'削除フラグ
	sLockFlag(0)  = Trim(rsd("LockFlag"))	'ロックフラグ
	iDualOpeNo(0) =      rsd("DualOpeNo")	'デュアル作業番号
	sGrpID(0)     = Trim(rsd("GroupID"))	'グループID
	sUsrID(0)     = Trim(rsd("UserID"))		'ユーザID
	sRDType(0)    = Trim(rsd("RecDel"))		'搬出入区分
	sContSize(0)  = Trim(rsd("ContSize"))	'コンテナサイズ
	bChkA(0)      = Trim(rsd("NotDelFlag"))	'搬出を載せない
	bChkB(0)      = Trim(rsd("NightFlag"))	'夕積シャーシ
	bChkC(0)      = Trim(rsd("MixSizeFlag"))'20/40兼用シャーシ
	sChID(0)      = Trim(rsd("ChassisId"))	'ｼｬｰｼID
	sPlace(0)     = Trim(rsd("Place"))		'場所
	AppDate(0)	  =      rsd("AppDate")		'申請日
	sTerminal(0)  = Trim(rsd("Terminal"))	'ターミナル
	rsd.Close

	If sDelFlag(0) = "Y" Then	'削除の場合
		sErrMsg = "交換相手が削除されています"
		Exit Function
	End If

	If sStatus(0) = "03" Then	'キャンセルの場合
		sErrMsg = "交換相手は作業がキャンセルされています"
		Exit Function
	End If
	If sStatus(0) = "04" Then	'完了の場合
		sErrMsg = "交換相手は作業を完了しています"
		Exit Function
	End If

	If sStatus(0) = "02" Then	'予約中の場合
		If sLockFlag(0) = "Y" Then	'ロック
			sErrMsg = "交換相手は作業が確定しています"
			Exit Function
		End If
	End If

	'申請情報の取得（指定作業番号）
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "対象が削除されています"
		Exit Function
	End If
	dWorkDate(1)  =      rsd("WorkDate")	'作業日
	sHH(1)    	  = Trim(rsd("Term"))		'時間帯
	iOpeOrder(1)  =      rsd("OpeOrder")	'作業順位
	sStatus(1)    = Trim(rsd("Status"))		'状態
	sDelFlag(1)   = Trim(rsd("DelFlag"))	'削除フラグ
	sLockFlag(1)  = Trim(rsd("LockFlag"))	'ロックフラグ
	iDualOpeNo(1) =      rsd("DualOpeNo")	'デュアル作業番号
	sGrpID(1)     = Trim(rsd("GroupID"))	'グループID
	sUsrID(1)     = Trim(rsd("UserID"))		'ユーザID
	sRDType(1)    = Trim(rsd("RecDel"))		'搬出入区分
	sContSize(1)  = Trim(rsd("ContSize"))	'コンテナサイズ
	bChkA(1)      = Trim(rsd("NotDelFlag"))	'搬出を載せない
	bChkB(1)      = Trim(rsd("NightFlag"))	'夕積シャーシ
	bChkC(1)      = Trim(rsd("MixSizeFlag"))'20/40兼用シャーシ
	sChID(1)      = Trim(rsd("ChassisId"))	'ｼｬｰｼID
	sPlace(1)     = Trim(rsd("Place"))		'場所
	AppDate(1)	  =      rsd("AppDate")		'申請日
	sTerminal(1)  = Trim(rsd("Terminal"))	'ターミナル
	rsd.Close

	If sDelFlag(1) = "Y" Then	'削除の場合
		sErrMsg = "対象が削除されています"
		Exit Function
	End If

	If sStatus(1) = "03" Then	'キャンセルの場合
		sErrMsg = "対象は作業がキャンセルされています"
		Exit Function
	End If
	If sStatus(1) = "04" Then	'完了の場合
		sErrMsg = "対象は作業を完了しています"
		Exit Function
	End If

	If sStatus(1) = "02" Then	'予約中の場合
		If sLockFlag(1) = "Y" Then	'ロック
			sErrMsg = "対象は作業が確定しています"
			Exit Function
		End If
	End If

	If sGrpID(0) <> sGrpID(1) Then
		sErrMsg = "交換相手のグループが異なります"
		Exit Function
	End If
'VP対応　搬入<->搬入または(搬出<->搬出,搬出<->空バン,空バン<->搬出,空バン<->空バン)
'''	If sRDType(0) <> sRDType(1) Then
'''		sErrMsg = "交換相手の搬出入区分が異なります"
'''		Exit Function
'''	End If
'搬入のチェック
	If (sRDType(0) = "R" and sRDType(1) <> "R") OR _
       (sRDType(0) <> "R" and sRDType(1) = "R")  Then
		sErrMsg = "交換相手の搬出入区分が異なります"
		Exit Function
	End If

	If sContSize(0) <> sContSize(1) Then
		sErrMsg = "交換相手のコンテナサイズが異なります"
		Exit Function
	End If
	If bChkA(0) <> bChkA(1) or _
	   bChkB(0) <> bChkB(1) or _
	   bChkC(0) <> bChkC(1) Then
		sErrMsg = "交換相手のシャーシ属性が異なります"
		Exit Function
	End If

'ターミナルのチェック
	If sTerminal(0) <> sTerminal(1) Then
		sErrMsg = "交換相手のターミナルが異なります"
		Exit Function
	End If

	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	rsd("UpdtTime")  = Now()				'更新日時
	rsd("UpdtPgCd")  = "SYWB031"			'更新ﾌﾟﾛｸﾞﾗﾑ名
	rsd("UpdtTmnl")  = "WEB"				'更新端末名
	rsd("WorkDate")  = dWorkDate(0)			'作業日
	rsd("Term")      = sHH(0)				'時間帯
	rsd("OpeOrder")  = iOpeOrder(0)			'作業順位
	rsd("HopeTerm")  = sHH(0)				'希望時間帯
	rsd("DualOpeNo") = iDualOpeNo(0)		'デュアル作業番号
	rsd("SendFlag")  = "Y"					'送信フラグ
'	rsd("AppDate") 	 = AppDate(0)			'申請日
	rsd.Update
	rsd.Close
	If iDualOpeNo(1) > 0 Then	'デュアル作業番号
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo(1))
		rsd("DualOpeNo") = CLng(sChgOpeNo)	'デュアル作業番号
		rsd("UpdtTime")  = Now()			'更新日時
		rsd("UpdtPgCd")  = "SYWB031"		'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl")  = "WEB"			'更新端末名
		rsd("SendFlag")  = "Y"				'送信フラグ
		rsd.Update
		rsd.Close
	End If
	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sChgOpeNo))
	rsd("UpdtTime")  = Now()				'更新日時
	rsd("UpdtPgCd")  = "SYWB031"			'更新ﾌﾟﾛｸﾞﾗﾑ名
	rsd("UpdtTmnl")  = "WEB"				'更新端末名
	rsd("WorkDate")  = dWorkDate(1)			'作業日
	rsd("Term")      = sHH(1)				'時間帯
	rsd("OpeOrder")  = iOpeOrder(1)			'作業順位
	rsd("HopeTerm")  = sHH(1)				'希望時間帯
	rsd("DualOpeNo") = iDualOpeNo(1)		'デュアル作業番号
	rsd("SendFlag")  = "Y"					'送信フラグ
'	rsd("AppDate") 	 = AppDate(1)			'申請日
	rsd.Update
	rsd.Close
	If iDualOpeNo(0) > 0 Then	'デュアル作業番号
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo(0))
		rsd("DualOpeNo") = CLng(sOpeNo)		'デュアル作業番号
		rsd("UpdtTime")  = Now()			'更新日時
		rsd("UpdtPgCd")  = "SYWB031"		'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl")  = "WEB"			'更新端末名
		rsd("SendFlag")  = "Y"				'送信フラグ
		rsd.Update
		rsd.Close
	End If
	sErrMsg = ""

End Function

%>
