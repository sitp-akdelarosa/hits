<% 

'予約移動処理
Function UpdOpeMov(conn, rsd, sOpeNo, sYMD, sHH, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sOpeNo			[ I ]作業番号
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]移動時間帯（HH）
'	sErrMsg			[ O ]エラーメッセージ

	Dim sStatus, sDelFlag, sLockFlag, iDualOpeNo, sGrpID, sUsrID
	Dim iNewDualOpeNo, iOpeOrder, sTerm_CD
	Dim sRDType, sContSize, bChkA, bChkB, bChkC, sChID, sPlace
'2001/03/26選択元・先の作業順位DUAL
	Dim iOpeOrder_M, iOpeOrder_S, iDualOpeNo_S, sDual

	sErrMsg = "ＤＢ更新エラー"

	'申請情報の取得（指定作業番号）
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "対象が削除されています"
		Exit Function
	End If
	sStatus    = Trim(rsd("Status"))			'状態
	sDelFlag   = Trim(rsd("DelFlag"))			'削除フラグ
	sLockFlag  = Trim(rsd("LockFlag"))			'ロックフラグ
	iDualOpeNo =      rsd("DualOpeNo")			'デュアル作業番号
	sGrpID     = Trim(rsd("GroupID"))			'グループID
	sUsrID     = Trim(rsd("UserID"))			'ユーザID
	sRDType    = Trim(rsd("RecDel"))			'搬出入区分
	sContSize  = Trim(rsd("ContSize"))			'コンテナサイズ
	bChkA      = Trim(rsd("NotDelFlag"))  = "Y"	'搬出を載せない
	bChkB      = Trim(rsd("NightFlag"))   = "Y"	'夕積シャーシ
	bChkC      = Trim(rsd("MixSizeFlag")) = "Y"	'20/40兼用シャーシ
	sChID      = Trim(rsd("ChassisId"))			'ｼｬｰｼID
	sPlace     = Trim(rsd("Place"))				'場所
	iOpeOrder_M =     rsd("OpeOrder")			'作業順位(選択された元)(3/26)
	sTerm_CD   = Trim(rsd("Terminal"))			'ターミナル
	rsd.Close

	'現在状態のチェック
	If sStatus = "04" Then	'完了の場合
		sErrMsg = "対象は作業を完了しています"
		Exit Function
	End If

	If sStatus = "02" Then	'予約中の場合
		If sLockFlag = "Y" Then	'ロック
			sErrMsg = "対象は作業が確定しています"
			Exit Function
		End If
	End If

	'搬入でシャーシが決定している（＝ＳＹ搬入済み）の場合
	If sRDType = "R" And sChID <> "" Then
		'シャーシの取得
		Call GetChassis(conn, rsd, sChID)
		If Trim(rsd("NotDelFlag")) = "Y" Then		'搬出を載せない
			bChkA = True
		End If
		If Trim(rsd("NightFlag")) = "Y" Then		'夕積シャーシ
			bChkB = True
		End If
		bChkC = Trim(rsd("MixSizeFlag")) = "Y"		'20/40兼用シャーシ
		rsd.Close
	End If

'2001/03/26	START	DUALの場合相手側に作業番号を見て作業順位が同じ場合はDUAL選択
	sDual = ""
	If iDualOpeNo > 0 Then 
		'申請情報の取得（指定作業番号）
		Call GetAppInfoOpeNo(conn, rsd, CLng(iDualOpeNo))
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "対象が削除されています"
			Exit Function
		End If
		iDualOpeNo_S =     rsd("DualOpeNo")		'デュアル作業番号(選択された相手)
		iOpeOrder_S  =     rsd("OpeOrder")		'作業順位(選択された相手)
		rsd.Close

'DUALOPE番号が交互に同じで作業順位が同じ場合DUAL選択
		If iOpeOrder_M = iOpeOrder_S And CLng(sOpeNo) = iDualOpeNo_S Then		
			sDual = "1"
		End If
	
	End iF
'2001/03/26	End

	'移動計算
	Call UpdOpeMovCalc(conn, rsd, sOpeNo, sGrpID, sUsrID, sYMD, sHH, _
						sRDType, _
						sContSize, bChkA, bChkB, bChkC,  _
						sChID, sPlace, sDual, sTerm_CD, _
						iNewDualOpeNo, iOpeOrder, sErrMsg)
	If sErrMsg <> "" Then
		Exit Function
	End If

	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	rsd("UpdtTime")  = Now()				'更新日時
	rsd("UpdtPgCd")  = "SYWB031"			'更新ﾌﾟﾛｸﾞﾗﾑ名
	rsd("UpdtTmnl")  = "WEB"				'更新端末名
	rsd("Status")    = "02"					'状態
	rsd("DelFlag")   = ""					'削除ﾌﾗｸ
	rsd("CReason")   = ""					'キャンセル理由
	rsd("Term")      = sHH					'時間帯
'2001/03/27DUAL識別が立っていて区分が搬入の時は作業順位をプラス１する
	If sDual =  "1" And  sRDType = "R" And _
		sHH <> "12" and sHH <> "A" and sHH <> "B" Then
			iOpeOrder = Cint(iOpeOrder) + 1		'作業順位計算
	End If
	rsd("OpeOrder")  = iOpeOrder			'作業順位

	rsd("HopeTerm")  = sHH					'希望時間帯
'2001/03/27DUAL識別が立たないもののみ更新	
	If sDual = "" Then						
		rsd("DualOpeNo") = iNewDualOpeNo	'デュアル作業番号
	End If
'2001/03/27DUAL識別が立たないもののみ更新	
	rsd("LockFlag")  = ""					'ロックフラグ
	rsd("SendFlag")  = "Y"					'送信フラグ
	rsd.Update
	rsd.Close

'2001/03/27DUAL識別が立たない場合は以下の処理
	If sDual = "" Then						
		If iDualOpeNo > 0 Then	'デュアル作業番号
			Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
			rsd("DualOpeNo") = 0				'デュアル作業番号
			rsd("UpdtTime")  = Now()			'更新日時
			rsd("UpdtPgCd")  = "SYWB031"		'更新ﾌﾟﾛｸﾞﾗﾑ名
			rsd("UpdtTmnl")  = "WEB"			'更新端末名
			rsd("SendFlag")  = "Y"				'送信フラグ
			rsd.Update
			rsd.Close
		End If
		If iNewDualOpeNo > 0 Then	'デュアル作業番号
			Call GetAppInfoOpeNoUpd(conn, rsd, iNewDualOpeNo)
			rsd("DualOpeNo") = CLng(sOpeNo)		'デュアル作業番号
			rsd("UpdtTime")  = Now()			'更新日時
			rsd("UpdtPgCd")  = "SYWB031"		'更新ﾌﾟﾛｸﾞﾗﾑ名
			rsd("UpdtTmnl")  = "WEB"			'更新端末名
			rsd("SendFlag")  = "Y"				'送信フラグ
			rsd.Update
			rsd.Close
		End If
	Else
'2001/03/27DUAL識別が立つ場合は以下の処理
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
		rsd("UpdtTime")  = Now()			'更新日時
		rsd("UpdtPgCd")  = "SYWB031"		'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl")  = "WEB"			'更新端末名
		rsd("Status")    = "02"				'状態
		rsd("DelFlag")   = ""				'削除ﾌﾗｸ
		rsd("CReason")   = ""				'キャンセル理由
		rsd("Term")      = sHH				'時間帯
		rsd("HopeTerm")  = sHH				'希望時間帯
		rsd("LockFlag")  = ""				'ロックフラグ
		rsd("SendFlag")  = "Y"				'送信フラグ
		rsd("OpeOrder")  = iOpeOrder		'作業順位
		rsd.Update
		rsd.Close
	End If

'test
'If sDual <> "" Then						
'sErrMsg = "テスト中 sDual= " & sDual & "作業順位=" & iOpeOrder
'Exit Function
'End If
'test

	sErrMsg = ""

End Function

'移動計算
Function UpdOpeMovCalc(conn, rsd, sOpeNo, sGrpID, sUsrID, sYMD, sHH, _
						sRDType, _
						sContSize, bChkA, bChkB, bChkC,  _
						sChID, sPlace, sDual, sTerm_CD, _
						iNewDualOpeNo, iOpeOrder, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sOpeNo			[ I ]作業番号
'	sGrpID     		[ I ]グループID
'	sUsrID     		[ I ]ユーザID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]移動時間帯（HH）
'	sRDType			[ I ]種類（D,R,V）
'	sContSize		[ I ]コンテナサイズ
'	bChkA			[ I ]搬入シャーシに搬出を載せない
'	bChkB			[ I ]搬入シャーシを夕積に使用
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sChID			[ I ]搬出シャーシID／搬入シャーシID
'	sPlace			[ I ]場所
'	sDual			[ I ]ＤＵＡＬ識別					3/27 add
'	sTerm_CD		[ I }対象ターミナル					10/4 add
'	iNewDualOpeNo	[ O ]デュアル作業番号
'	iOpeOrder		[ O ]作業順位
'	sErrMsg			[ O ]エラーメッセージ

	Dim sWk, iAssignCnt, bOk, iCnt
	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime

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
		sWk = TermNumToSym(iNextTime)
		If sHH < sWk or (sHH = sWk and iNextStat = 2) Then
			sErrMsg = "指定時間帯の予約は締め切られています"
			Exit Function
		End If
		If sHH <> TermNumToSym(iNextTime) Then
			iOpenSlot = 0	'開放枠数
		End If
	Else
		iOpenSlot = 0	'開放枠数
	End If

	'時間帯指定の場合にはグループの割り当て枠を取得
	iAssignCnt = -1
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'グループの割り当て枠を取得
		iAssignCnt = GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)
	End If

	'時間帯指定の場合は枠をチェック
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'予約枠によるチェック
		If Not AppUpdCheckTerm(conn, rsd, _
						sUsrID, sGrpID, _
						sYMD, sHH, _
						sRDType, _
						sContSize, bChkA, bChkB, bChkC,  _
						sChID, sPlace, sDual, sTerm_CD, _
						iAssignCnt, iOpenSlot, iNewDualOpeNo, iOpeOrder, sErrMsg) Then
			Exit Function
		End If
	Else
		iNewDualOpeNo = 0		'デュアル作業番号
		iOpeOrder = 0			'作業順位
	End If

End Function

'予約枠によるチェック
Function AppUpdCheckTerm(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContSize, bChkA, bChkB, bChkC,  _
			sChID, sPlace, sDual, sTerm_CD, _
			iAssignCnt, iOpenSlot, iNewDualOpeNo, iOpeOrder, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]予約時間帯（HH）
'	sRDType			[ I ]種類（D,R,V）
'	sContSize		[ I ]コンテナサイズ
'	bChkA			[ I ]搬入シャーシに搬出を載せない
'	bChkB			[ I ]搬入シャーシを夕積に使用
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sChID			[ I ]搬出シャーシID／搬入シャーシID
'	sPlace			[ I ]場所
'	sDual			[ I ]DUAL識別		'2001/03/27
'	sTerm_CD		[ I ]対象ターミナル '2001/10/04
'	iAssignCnt 		[ I ]グループ割当数
'	iOpenSlot 		[ I ]開放枠数
'	iNewDualOpeNo	[ O ]デュアル作業番号
'	iOpeOrder		[ O ]作業順位
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	Dim iCnt	'現在往復回数（＝登録グループ枠数）

	AppUpdCheckTerm = False

	'申請情報の取得（指定日、指定時間帯、指定グループ）
	Call GetAppInfoTerm(conn, rsd, sYMD, sHH, sGrpID)
	iCnt = 0	'現在往復回数（＝登録グループ枠数）
	iNewDualOpeNo = 0		'デュアル作業番号
	iOpeOrder = 0		'作業順位
	Do Until rsd.EOF
'2001/03/27 DUAL識別が立ってない場合のみ以下の処理をする
		If sDual = "" Then
			'搬入でデュアルになれる場合(選択されたのが搬入で・該当の搬出相手を捜す)
			If sRDType = "R" and (Not bChkA) and (Not bChkB) Then
									'ペアになり得る搬出かチェック
					'該当する相手が、搬出かまたは空バンの時に変更する
					'''				If rsd("DualOpeNo") = 0 and _
					'''VP対応		   rsd("RecDel") = "D" and _
					'''				   (Trim(rsd("ChassisId")) = "" or _
					'''				    Trim(rsd("ChassisId")) = sChID) Then
					'If rsd("DualOpeNo") = 0 and _
					'   rsd("RecDel") <> "R" and _			
					'   (Trim(rsd("ChassisId")) = "" or _
					'    Trim(rsd("ChassisId")) = sChID) Then
					'	'搬入が香椎・箱崎かまたは、それ以外の前受けの場合は同じVPターミナルの場合
					'	If (sTerm_CD = "KA" OR sTerm_CD = "HK") or _
					'	   (sTerm_CD <> "KA" AND sTerm_CD <> "HK" AND sTerm_CD = Trim(rsd("Terminal")))	  Then	
					'		'兼用かサイズがマッチする場合
					'		If bChkC or _
					'		   (rsd("ContSize") =  "20" and sContSize =  "20") or _
					'		   (rsd("ContSize") <> "20" and sContSize <> "20") Then
					'			iNewDualOpeNo = rsd("OpeNo")	'デュアル作業番号
					'			iOpeOrder = rsd("OpeOrder") - 1
					'		End If
					'	End If
					'End If
				'02/04/27 (DUALは同一ターミナルにする) sta
				If rsd("DualOpeNo") = 0 and _
				   rsd("RecDel") <> "R" and _			
				   (Trim(rsd("ChassisId")) = "" or _
				    Trim(rsd("ChassisId")) = sChID) Then
					'同じVPターミナル
					If sTerm_CD = Trim(rsd("Terminal"))	  Then	
						'兼用かサイズがマッチする場合
						If bChkC or _
						   (rsd("ContSize") =  "20" and sContSize =  "20") or _
						   (rsd("ContSize") <> "20" and sContSize <> "20") Then
							iNewDualOpeNo = rsd("OpeNo")	'デュアル作業番号
							iOpeOrder = rsd("OpeOrder") - 1
						End If
					End If
				End If
				'02/04/27 (DUALは同一ターミナルにする) end
			End If

			'搬出でデュアルになれる場合(選択されたのが搬出で・該当の搬入相手を捜す)
			'搬入以外でデュアルになれる場合(選択されたのが搬入以外で・該当の搬入相手を捜す)
'''			If sRDType = "D" Then	
			If sRDType <> "R" Then			'搬入でない
				'ペアになり得る搬入かチェック
				If rsd("DualOpeNo") = 0 and _
				   rsd("RecDel") = "R" and _
				   rsd("NotDelFlag") <> "Y" and _
				   rsd("NightFlag") <> "Y" and _
				   (sChID = "" or _
				    Trim(rsd("ChassisId")) = sChID) Then
						'同じターミナルに限る
						If sTerm_CD = Trim(rsd("Terminal")) Then	
							'兼用かサイズがマッチする場合
							If rsd("MixSizeFlag") = "Y" or _
							   (rsd("ContSize") =  "20" and sContSize =  "20") or _
							   (rsd("ContSize") <> "20" and sContSize <> "20") Then
								iNewDualOpeNo = rsd("OpeNo")	'デュアル作業番号
								iOpeOrder = rsd("OpeOrder") + 1
							End If
						End If
				End If
			End If
		End If

'''		If rsd("DualOpeNo") = 0 or _
'''		   rsd("RecDel") = "D" Then
'''		   rsd("RecDel") <> "R" Then	'02/04/27 sywb012.incに合わせる
		If rsd("DualOpeNo") = 0 or rsd("RecDel") <> "R" Then
			'搬出は単独のみカウント＝往復回数のカウント
			iCnt = iCnt + 1	'現在往復回数（＝登録グループ枠数）
		End If

		rsd.MoveNext
	Loop
	rsd.Close

	If iNewDualOpeNo = 0 and _
	   iOpenSlot = 0 and _
	   iCnt >= iAssignCnt Then
		sErrMsg = "割り当て枠がありません"
		Exit Function
	End If

	'他の作業とデュアルにならない場合は作業順位を計算
	'またはDUAL識別フラグが立っている場合は順位を計算する
	If iOpeOrder = 0 Or sDual <> "" Then
		'新規作業順位の取得（指定日、指定時間帯）
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, sRDType)
	End If

	AppUpdCheckTerm = True

End Function

%>
