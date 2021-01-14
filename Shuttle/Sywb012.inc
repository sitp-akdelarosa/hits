<% 

'予約更新処理	2003/08/27(ICCT 認証IDに追加)
Function UpdateApp(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType(), _
			sContNoRec(), sBKNo(), sContSizeRec(), bChkA(), bChkB(), bChkC(),  _
			sContNoDel(), sChID(), sBLNo(), sContSizeDel(), sDeliverTo(), sReceiveFrom(),  _
			sTerm_CD, sNinID(), sErrMsg(), sOpeNoRec(), sOpeNoDel()) 
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]予約時間帯（HH）
'	sRDType			[ I ]種類（DUAL,REC,DEL,空白）
'	sContNoRec 		[ I ]搬入コンテナ番号
'	sBKNo			[ I ]搬入ブッキング番号
'	sContSizeRec	[ I ]搬入コンテナサイズ
'	bChkA			[ I ]搬入シャーシに搬出を載せない
'	bChkB			[ I ]搬入シャーシを夕積に使用
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sContNoDel 		[ I ]搬出コンテナ番号
'	sChID			[ I ]搬出シャーシID
'	sBLNo 			[ I ]搬出ＢＬ番号
'	sContSizeDel	[I/O]搬出コンテナサイズ
'	sDeliverTo		[ I ]搬出先
'	sReceiveFrom	[ I ]搬入元
'	sTerm_CD		[ I ]ターミナル
'	sNinID			[ I ]認証ＩＤ
'	sErrMsg			[ O ]エラーメッセージ(0:ALL,1～4:各予約)（成功の場合は空白）
'	sOpeNoRec		[ O ]搬入予約番号
'	sOpeNoDel		[ O ]搬出予約番号

	Dim i, sWk, iAssignCnt, bOk, iDualOpeNo, iOpeOrder, iCnt
	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime
	Dim sVslCode, sVoyage, sOpenFlag
	Dim sOpeRec, sOpeDel

	sErrMsg(0) = ""

'I20040206 S
    Dim sAn,ilen
    sAn = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'I20040206 E

	'*** 予約日、予約時間帯が有効かチェック ***
	'現在日付取得
	sDateNow = GetYMDStr(Date())

	If sDateNow > sYMD Then
		sErrMsg(0) = "作業日が適切ではありません"
		Exit Function
	End If
	If sDateNow = sYMD Then	'現在日指定の場合
		'運行状況を取得
		Call GetOpeStatus(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime)
		sWk = TermNumToSym(iNextTime)
		If sHH < sWk or (sHH = sWk and iNextStat = 2) Then
			sErrMsg(0) = "指定時間帯の予約は締め切られています"
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
	sOpenFlag = ""
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'グループの割り当て枠を取得
		iAssignCnt = GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)

		'グループ時間帯情報の取得
		sOpenFlag = GetGrpSlotHH(conn, rsd, sGrpID, sYMD, sHH)

	End If

	'*** 予約１～４を登録 ***
	For i = 0 To 3

		If sRDType(i) <> "" Then

			'予約チェック 2003/08/27 認証ID追加
			bOk = AppUpdCheck(conn, rsd, _
					sUsrID, sGrpID, _
					sYMD, sHH, _
					sRDType(i), _
					sContNoRec(i), sBKNo(i), sContSizeRec(i), bChkA(i), bChkB(i), bChkC(i),  _
					sContNoDel(i), sChID(i), sBLNo(i), sVslCode, sVoyage, sContSizeDel(i), _
					sOpeRec, sOpeDel, sTerm_CD, sNinID(i), _
					iAssignCnt, iOpenSlot, _
					iDualOpeNo, iOpeOrder, iCnt, _
					sErrMsg(i + 1)) 

			If bOk Then
				'予約処理
				Call UpdateApp2(conn, rsd, _
					sUsrID, sGrpID, _
					sYMD, sHH, _
					iAssignCnt, sOpenFlag, _
					iDualOpeNo, iOpeOrder, iCnt, _
					sRDType(i), _
					sContNoRec(i), sBKNo(i), sContSizeRec(i), bChkA(i), bChkB(i), bChkC(i),  _
					sContNoDel(i), sChID(i), sBLNo(i), sVslCode, sVoyage, sContSizeDel(i), sDeliverTo(i), _
					sOpeRec, sOpeDel, sReceiveFrom(i), sTerm_CD, _
					sErrMsg(i + 1), sOpeNoRec(i), sOpeNoDel(i)) 
			End If

'I20040206 S
			'搬入コンテナ番号の英数字チェック
		    For ilen = 1 To len(sContNoRec(i))
		        If InStrRev(sAn, Mid(sContNoRec(i), ilen, 1)) < 1 Then
					sErrMsg(i + 1) = "コンテナ番号に英数字以外の文字が含まれています"
		        End If
		    Next
'I20040206 E


		End If
	Next

End Function

'予約更新処理２
Function UpdateApp2(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			iAssignCnt, sOpenFlag, _
			iDualOpeNo, iOpeOrder, iCnt, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC, _
			sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, sDeliverTo, _
			sOpeRec, sOpeDel, sReceiveFrom, sTerm_CD, _
			sErrMsg, sOpeNoRec, sOpeNoDel) 
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	iAssignCnt 		[ I ]グループ割当数
'	sOpenFlag		[ I ]開放フラグ
'	iDualOpeNo		[ I ]デュアル作業番号
'	iOpeOrder		[ I ]作業順位
'	iCnt			[ I ]予約枠数
'	sRDType			[ I ]種類（DUAL,REC,DEL,空白）
'	sContNoRec 		[ I ]搬入コンテナ番号
'	sBKNo			[ I ]搬入ブッキング番号
'	sContSizeRec		[ I ]搬入コンテナサイズ
'	bChkA			[ I ]搬入シャーシに搬出を載せない
'	bChkB			[ I ]搬入シャーシを夕積に使用
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sContNoDel 		[ I ]搬出コンテナ番号
'	sChID			[ I ]搬出シャーシID
'	sBLNo 			[ I ]搬出ＢＬ番号
'	sVslCode		[ I ]搬出船名（ＢＬ指定時）
'	sVoyage			[ I ]搬出次航（ＢＬ指定時）
'	sContSizeDel		[ I ]搬出コンテナサイズ
'	sDeliverTo		[ I ]搬出先
'	sOpeRec			[ I ]搬入オペレータ
'	sOpeDel			[ I ]搬出オペレータ
'	sReceiveFrom		[ I ]搬入元
'	sTerm_CD		[ I ]ターミナル
'	sErrMsg			[ O ]エラーメッセージ（成功の場合は空白）
'	sOpeNoRec		[ O ]搬入予約番号
'	sOpeNoDel		[ O ]搬出予約番号


	Dim sql, bTermFlag

'I20040209 S
	Dim sYLocExt
	sYLocExt = ""
'I20040209 E

	sErrMsg = "ＤＢ更新エラーが発生しました"

	sOpeNoRec = "0"
	sOpeNoDel = "0"

	If sRDType = "REC" or sRDType = "DUAL" Then
		'作業番号の発番
		Call GetNewOpeNo(conn, rsd, sYMD, sOpeNoRec)
	End If
	If sRDType = "DEL" or sRDType = "DUAL" Then
		'作業番号の発番
		Call GetNewOpeNo(conn, rsd, sYMD, sOpeNoDel)
'I20040209 S
		Call GetOutCont(conn, rsd, sContNoDel)
		If rsd.EOF = False Then
			sYLocExt = rsd("YLocExt")
		End If
		rsd.Close
'I20040209 E
	End If

	'時間帯指定の場合
	bTermFlag = False
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		bTermFlag = True
		'開放枠を使用する場合はこれを減算
		If iDualOpeNo = 0 and (sOpenFlag = "Y" or iAssignCnt <= iCnt) Then
			'開放枠の減算
			Call SubmitOpenSlot(conn, rsd)
		End If
	End If

	'申請情報検索
	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iDualOpeNo)
	rsd.Open sql, conn, 0, 2, 1

	'他の申請作業とデュアルになる場合
	If iDualOpeNo > 0 Then
		rsd("UpdtTime") 	= Now()				'更新日時
		rsd("UpdtPgCd")		= "SYWB012"			'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl") 	= "WEB"				'更新端末名
		If sOpeNoRec <> "0" Then
			rsd("DualOpeNo") = CLng(sOpeNoRec)	'デュアル作業番号
		Else
			rsd("DualOpeNo") = CLng(sOpeNoDel)	'デュアル作業番号
		End If
		rsd("SendFlag")    	= "Y"				'変更送信ﾌﾗｸ
	End If

	If sRDType = "REC" or sRDType = "DUAL" Then
		rsd.AddNew
		rsd("OpeNo")       = sOpeNoRec			'作業番号
		rsd("UpdtTime")    = Now()				'更新日時
		rsd("UpdtPgCd")    = "SYWB012"			'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl")    = "WEB"				'更新端末名
		rsd("UserID")      = sUsrID				'ユーザｺｰﾄ
		rsd("GroupID")     = sGrpID				'ｸﾞﾙｰﾌﾟｺｰﾄ
		rsd("AppDate")     = Now()				'申請日
		rsd("WorkDate")    = ChgYMDDate(sYMD)	'作業日
		rsd("RecDel")      = "R"				'作業種類
		rsd("ContNo")      = sContNoRec			'ｺﾝﾃﾅ番号
		rsd("BookNo")      = sBKNo				'ﾌﾞｯｷﾝｸﾞ番号
		rsd("Status")      = "02"				'状態
		rsd("Place")       = ""					'場所
		rsd("ChassisId")   = ""					'ｼｬｰｼID
		rsd("DelFlag")     = ""					'削除ﾌﾗｸ
		rsd("CReason")     = ""					'キャンセル理由
		rsd("WorkFlag")    = ""					'ｼｬﾄﾙ作業中ﾌﾗｸ
		rsd("SendFlag")    = "Y"				'変更送信ﾌﾗｸ
		If sContSizeRec = "20" Then
			rsd("Size20Flag")  = "Y"			'20フィートフラグ
		Else
			rsd("Size20Flag")  = ""				'20フィートフラグ
		End If
'搬入時も船名・次航をセットする(02/02/09)
'''		rsd("VslCode")     = ""					'船名ｺｰﾄ
'''		rsd("Voyage")      = ""					'次航
		rsd("VslCode")     = sVslCode				'船名ｺｰﾄ
		rsd("Voyage")      = sVoyage				'次航
		rsd("BLNo")        = ""					'ＢＬ番号
		rsd("Term")        = sHH				'時間帯
		If iDualOpeNo > 0 Then	'他の作業とデュアルになる場合
			rsd("OpeOrder") = iOpeOrder			'作業順位
		Else
			rsd("OpeOrder") = iOpeOrder			'作業順位
		End If
		rsd("HopeTerm")    = sHH				'希望時間帯
		If iDualOpeNo > 0 Then	'他の作業とデュアルになる場合
			rsd("DualOpeNo") = iDualOpeNo		'デュアル作業番号
		Else
'			If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
'				rsd("DualOpeNo") = CLng(sOpeNoDel)	'デュアル作業番号
'			Else
'				rsd("DualOpeNo") = 0				'デュアル作業番号
'			End If
			rsd("DualOpeNo") = CLng(sOpeNoDel)	'デュアル作業番号
		End If
		rsd("ContSize")    = sContSizeRec		'コンテナサイズ
		If bChkA Then
			rsd("NotDelFlag") = "Y"				'搬出を載せない
		Else
			rsd("NotDelFlag") = ""				'搬出を載せない
		End If
		If bChkB Then
			rsd("NightFlag") = "Y"				'夕積シャーシ
		Else
			rsd("NightFlag") = ""				'夕積シャーシ
		End If
		If bChkC Then
			rsd("MixSizeFlag") = "Y"			'20/40兼用シャーシ
		Else
			rsd("MixSizeFlag") = ""				'20/40兼用シャーシ
		End If
		rsd("DeliverTo")   = ""					'搬出先
		rsd("LockFlag")    = ""					'ロックフラグ
		rsd("Operator")    = sOpeRec			'オペレータコード
		rsd("OpeBookNo")   = ""					'自社ブッキング番号
		rsd("ReservNo")    = 0					'予約番号

		rsd("RecContSize") = sContSizeRec		'搬入ｺﾝﾃﾅｻｲｽﾞ
		rsd("ContHeight")  = ""					'搬入ｺﾝﾃﾅﾊｲﾄ
		rsd("ContType")    = ""					'搬入ｺﾝﾃﾅﾀｲﾌﾟ
		rsd("Material")    = ""					'搬入ｺﾝﾃﾅ材質
		rsd("TareWeight")  = 0					'搬入ｺﾝﾃﾅﾃｱｳｪｲﾄ
		rsd("Gweight")     = 0					'搬入ｺﾝﾃﾅｸﾞﾛｽｳｪｲﾄ
		rsd("SoundUnsound")  = ""				'搬入Sound/Unsound
		rsd("Trucker")  = ""					'搬入運送業者
		rsd("ReceiveFrom")  = sReceiveFrom		'搬入元
		rsd("CustOK")  = ""						'搬入ｺﾝﾃﾅ通関/未通関
		rsd("Terminal")  = sTerm_CD				'該当ターミナル(VP予約)暫定修正
		rsd("VPBookNo")      = ""				'搬出ﾌﾞｯｷﾝｸﾞ番号(VP予約)
		rsd("LineCode")      = ""				'航路コード(VP予約)
		rsd("HeadID")        = ""				'ヘッドＩＤ(ICCT対応)03/07/31
		rsd("YLocExt")       = sYLocExt			'外部座標 I20040209
	End If
	If sRDType = "DEL" or sRDType = "DUAL" Then
		rsd.AddNew
		rsd("OpeNo")       = sOpeNoDel			'作業番号
		rsd("UpdtTime")    = Now()				'更新日時
		rsd("UpdtPgCd")    = "SYWB012"			'更新ﾌﾟﾛｸﾞﾗﾑ名
		rsd("UpdtTmnl")    = "WEB"				'更新端末名
		rsd("UserID")      = sUsrID				'ユーザｺｰﾄ
		rsd("GroupID")     = sGrpID				'ｸﾞﾙｰﾌﾟｺｰﾄ
		rsd("AppDate")     = Now()				'申請日
		rsd("WorkDate")    = ChgYMDDate(sYMD)	'作業日
		rsd("RecDel")      = "D"				'作業種類
		rsd("ContNo")      = sContNoDel			'ｺﾝﾃﾅ番号
		rsd("BookNo")      = ""					'ﾌﾞｯｷﾝｸﾞ番号
		rsd("Status")      = "02"				'状態
		rsd("Place")       = "CY"				'場所
		rsd("ChassisId")   = sChID				'ｼｬｰｼID
		rsd("DelFlag")     = ""					'削除ﾌﾗｸ
		rsd("CReason")    = ""					'キャンセル理由
		rsd("WorkFlag")    = ""					'ｼｬﾄﾙ作業中ﾌﾗｸ
		rsd("SendFlag")    = "Y"				'変更送信ﾌﾗｸ
		If sContSizeDel = "20" Then
			rsd("Size20Flag")  = "Y"			'20フィートフラグ
		Else
			rsd("Size20Flag")  = ""				'20フィートフラグ
		End If
		rsd("VslCode")     = sVslCode			'船名ｺｰﾄ
		rsd("Voyage")      = sVoyage			'次航
		rsd("BLNo")        = sBLNo				'ＢＬ番号
		rsd("Term")        = sHH				'時間帯
		If Not bTermFlag Then		'時間帯指定でない場合
			rsd("OpeOrder") = 0					'作業順位
		ElseIf iDualOpeNo > 0 Then	'他の作業とデュアルになる場合
			rsd("OpeOrder") = iOpeOrder			'作業順位
		Else
			rsd("OpeOrder") = iOpeOrder			'作業順位
		End If
		rsd("HopeTerm")    = sHH				'希望時間帯
		If iDualOpeNo > 0 Then	'他の作業とデュアルになる場合
			rsd("DualOpeNo") = iDualOpeNo		'デュアル作業番号
		Else
'			If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
'				rsd("DualOpeNo") = CLng(sOpeNoRec)	'デュアル作業番号
'			Else
'				rsd("DualOpeNo") = 0				'デュアル作業番号
'			End If
			rsd("DualOpeNo") = CLng(sOpeNoRec)	'デュアル作業番号
		End If
		rsd("ContSize")    = sContSizeDel		'コンテナサイズ
		rsd("NotDelFlag")  = ""					'搬出を載せない
		rsd("NightFlag")   = ""					'夕積シャーシ
		rsd("MixSizeFlag") = ""					'20/40兼用シャーシ
		rsd("DeliverTo")   = sDeliverTo			'搬出先
		rsd("LockFlag")    = ""					'ロックフラグ
		rsd("Operator")    = sOpeDel			'オペレータコード
		rsd("OpeBookNo")   = ""					'自社ブッキング番号
		rsd("ReservNo")    = 0					'予約番号

		rsd("RecContSize") = sContSizeDel		'搬入ｺﾝﾃﾅｻｲｽﾞ
		rsd("ContHeight")  = ""					'搬入ｺﾝﾃﾅﾊｲﾄ
		rsd("ContType")    = ""					'搬入ｺﾝﾃﾅﾀｲﾌﾟ
		rsd("Material")    = ""					'搬入ｺﾝﾃﾅ材質
		rsd("TareWeight")  = 0					'搬入ｺﾝﾃﾅﾃｱｳｪｲﾄ
		rsd("Gweight")     = 0					'搬入ｺﾝﾃﾅｸﾞﾛｽｳｪｲﾄ
		rsd("SoundUnsound")  = ""				'搬入Sound/Unsound
		rsd("Trucker")  = ""					'搬入運送業者
		rsd("ReceiveFrom")  = ""				'搬入元
		rsd("CustOK")  = ""						'搬入ｺﾝﾃﾅ通関/未通関
		rsd("Terminal")  = sTerm_CD				'該当ターミナル(VP予約)
		rsd("VPBookNo")      = ""				'搬出ﾌﾞｯｷﾝｸﾞ番号(VP予約)
		rsd("LineCode")      = ""				'航路コード(VP予約)
		rsd("HeadID")        = ""				'ヘッドＩＤ(ICCT対応)03/07/31
		rsd("YLocExt")       = sYLocExt			'外部座標 I20040209
	End If

	rsd.Update
	rsd.Close
	sErrMsg = ""

End Function

'予約チェック
Function AppUpdCheck(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
			sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, _
			sOpeRec, sOpeDel, sTerm_CD, sNinID, _
			iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, iCnt, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]予約時間帯（HH）
'	sRDType			[ I ]種類（DUAL,REC,DEL,空白）
'	sContNoRec 		[ I ]搬入コンテナ番号
'	sBKNo			[ I ]搬入ブッキング番号
'	sContSizeRec	[ I ]搬入コンテナサイズ
'	bChkA			[ I ]搬入シャーシに搬出を載せない
'	bChkB			[ I ]搬入シャーシを夕積に使用
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sContNoDel 		[ I ]搬出コンテナ番号
'	sChID			[ I ]搬出シャーシID
'	sBLNo 			[ I ]搬出ＢＬ番号
'	sVslCode		[ O ]搬出船名（ＢＬ指定時）
'	sVoyage			[ O ]搬出次航（ＢＬ指定時）
'	sContSizeDel	[I/O]搬出コンテナサイズ
'	sOpeRec			[ O ]搬入オペレータ
'	sOpeDel			[ O ]搬出オペレータ
'	sTerm_CD		[ I ]ターミナル
'   sNinID			[ I ]認証ＩＤ
'	iAssignCnt 		[ I ]グループ割当数
'	iOpenSlot 		[ I ]開放枠数
'	iDualOpeNo		[ O ]デュアル作業番号
'	iOpeOrder		[ O ]作業順位
'	iCnt			[ O ]予約枠数
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheck = False

	'搬出時のチェック(2003/08/27 ﾊﾟﾗﾒｰﾀ追加 sNinID ICCT対応)
	If sRDType = "DEL" or sRDType = "DUAL" Then
		If Not AppUpdCheckDel(conn, rsd, _
						sGrpID, _
						sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, sOpeDel, _
						bChkC, sTerm_CD, sNinID, sErrMsg) Then
			Exit Function
		End If
	End If

	'搬入時のチェック(02/02/09 ﾊﾟﾗﾒｰﾀ追加 sVsl, sVoy)
	If sRDType = "REC" or sRDType = "DUAL" Then
		If Not AppUpdCheckRec(conn, rsd, _
						sContNoRec, sBKNo, sContSizeRec, sOpeRec,  _
						sVslCode, sVoyage, sErrMsg) Then
			Exit Function
		End If
	End If

	'搬出入時のチェック
	If sRDType = "DUAL" Then
		'兼用シャーシでない場合はサイズのマッチをチェック(このコメントの対応をする2001/6/2 M.OKIMI)
		If (Not bChkC) Then
			If ((sContSizeRec =  "20" and sContSizeDel <> "20") or _
				(sContSizeRec <> "20" and sContSizeDel =  "20")) Then
				sErrMsg = "搬出入のコンテナサイズが異なります"
				Exit Function
			End If
		End If
	End If

	'時間帯指定の場合は枠をチェック
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'予約枠によるチェック
		If Not AppUpdCheckTerm(conn, rsd, _
						sUsrID, sGrpID, _
						sYMD, sHH, _
						sRDType, _
						sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
						sContNoDel, sChID, sBLNo, sContSizeDel, _
						iAssignCnt, iOpenSlot, sTerm_CD, _
						iDualOpeNo, iOpeOrder, iCnt, sErrMsg) Then

			Exit Function
		End If
	Else
		iDualOpeNo = 0		'デュアル作業番号
		iOpeOrder = 0		'作業順位
		iCnt = 0			'予約枠数
	End If

	AppUpdCheck = True
End Function

'搬出時予約チェック
Function AppUpdCheckDel(conn, rsd, _
			sGrpID, _
			sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, sOpeDel, _
			bChkC, sTerm_CD, sNinID, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sGrpID			[ I ]グループID
'	sContNoDel 		[ I ]搬出コンテナ番号
'	sChID			[ I ]搬出シャーシID
'	sBLNo 			[ I ]搬出ＢＬ番号
'	sVslCode		[ O ]搬出船名（ＢＬ指定時）
'	sVoyage			[ O ]搬出次航（ＢＬ指定時）
'	sContSizeDel	[I/O]搬出コンテナサイズ
'	sOpeDel			[ O ]搬出オペレータ
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sTerm_CD  		[ I ]ターミナル
'   sNinID			[ I ]認証ID
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheckDel = False

	sVslCode = ""
	sVoyage  = ""

	If sContNoDel <> "" Then	'コンテナ指定
		'申請の有無チェック（指定コンテナ）
		If ChkAppInfoExt(conn, rsd, sContNoDel) Then
			sErrMsg = "既に予約済みのコンテナです（" & sContNoDel & "）"
			Exit Function
		End If

		'搬出許可コンテナの取得
		Call GetOutCont(conn, rsd, sContNoDel)
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "搬出コンテナが在庫していません（" & sContNoDel & "）"
			Exit Function
		End IF
'VP対応
'''		If  Trim(rsd("Terminal")) <> "KA" Then
'''			rsd.Close
'''			sErrMsg = "搬出コンテナが香椎に在庫していません（" & sContNoDel & "）"
'''			Exit Function
'''		End If

		If  sTerm_CD = "KA" Then
			If Trim(rsd("Terminal")) <> "KA" Then
				rsd.Close
				sErrMsg = "搬出コンテナが香椎に在庫していません（" & sContNoDel & "）"
				Exit Function
			End If
		End If
'<2003/08/25 ICCT対応>
'''		If  sTerm_CD = "HK" Then
'''			If  Trim(rsd("Terminal")) <> "HK" Then
'''				rsd.Close
'''				sErrMsg = "搬出コンテナが箱崎に在庫していません（" & sContNoDel & "）"
'''				Exit Function
'''			End If
'''		End If

   		If  sTerm_CD = "IC" Then
   			If  Trim(rsd("Terminal")) <> "IC" Then
   				rsd.Close
   				sErrMsg = "搬出コンテナがＩＣＣＴに在庫していません（" & sContNoDel & "）"
   				Exit Function
   			End If
		End If

		If  Trim(rsd("FullEmpty")) = "F" Then
			If IsNull(rsd("DelOKDate")) Then
				rsd.Close
				sErrMsg = "指定コンテナは搬出不可です（" & sContNoDel & "）"
				Exit Function
			End If
			If sYMD < GetYMDStr(rsd("DelOKDate")) Then
				rsd.Close
				sErrMsg = "指定コンテナは搬出不可です（" & sContNoDel & "）"
				Exit Function
			End If
			If not IsNull(rsd("DemFTDate")) Then
				If sYMD > GetYMDStr(rsd("DemFTDate")) Then
					rsd.Close
					sErrMsg = "指定コンテナは搬出不可です（" & sContNoDel & "）"
					Exit Function
				End If
			End If
			If not IsNull(rsd("OLTFrom"))  Then	
				If GetYMDStr(rsd("OLTFrom")) > sYMD or _
				   GetYMDStr(rsd("OLTTo"))   < sYMD Then
					rsd.Close
					sErrMsg = "指定コンテナは搬出不可です（" & sContNoDel & "）"
					Exit Function
				End If
			End If
		End If
'リーファーコンテナを搬出可能にする(01/09/04)
'''		If  Left( Trim(rsd("RHO")), 1) = "R" Or _
'''			Right(Trim(rsd("RHO")), 1) = "R" Then

'''			rsd.Close
'''			sErrMsg = "リーファーコンテナの為不可です（" & sContNoDel & "）"
'''			Exit Function
'''		End If

'認証ＩＤに値がある場合は入力した値と等しくないとエラー(2003/08/27)

		If  Trim(rsd("DONo")) <> "" Then
			If  Trim(rsd("DONo")) <> sNinID  Then
   				rsd.Close
   				sErrMsg = "認証ＩＤを正しく入力してください（" & sContNoDel & "）"
   				Exit Function
			End If
   		End If

		sContSizeDel = Trim(rsd("ContSize"))
		sOpeDel = Trim(rsd("Operator"))
'02/02/09 ADD
		sVslCode = Trim(rsd("VslCode"))
		sVoyage  = Trim(rsd("Voyage"))
'02/02/09 ADD
		rsd.Close
	Else						'ＢＬ指定
		'搬出許可マニフェストコンテナの取得(複数ＣＹ対応(01/10/01)
		Call GetOutBLContCont(conn, rsd, sBLNo, sTerm_CD)
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "ＢＬ番号が正しくありません（" & sBLNo & "）"
			Exit Function
		End IF
		sVslCode = Trim(rsd("VslCode"))
		sVoyage  = Trim(rsd("Voyage"))
		sOpeDel = Trim(rsd("Operator"))

		'重複チェック
		Do Until rsd.EOF
			If sVslCode <> Trim(rsd("VslCode")) or _
			   sVoyage  <> Trim(rsd("Voyage")) Then
				rsd.Close
				sErrMsg = "同一のＢＬ番号が存在しますのでコンテナ番号指定で予約して下さい"
				Exit Function
			End If
			rsd.MoveNext
		Loop

		rsd.Close
	End If

	'シャーシをチェック
	If sChID <> "" Then	'搬出シャーシID
		'シャーシの取得
	 	Call GetChassis(conn, rsd, sChID)
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "シャーシが登録されていません（" & sChID & "）"
			Exit Function
		End IF
		If Trim(rsd("GroupID")) <> sGrpID Then
			rsd.Close
			sErrMsg = "シャーシの所属が異なります（" & sChID & "）"
			Exit Function
		End If
		If rsd("MixSizeFlag") <> "Y" Then	'兼用シャーシでない
			If bChkC Then
				rsd.Close
				sErrMsg = "20/40兼用シャーシではありません（" & sChID & "）"
				Exit Function
			End If
			If (sContSizeDel <> "20" and rsd("Size20Flag") =  "Y") or _
			   (sContSizeDel =  "20" and rsd("Size20Flag") <> "Y") Then
				rsd.Close
				sErrMsg = "シャーシサイズが異なります（" & sChID & "）"
				Exit Function
			End If
		End If
		rsd.Close

		'指定シャーシが同一作業日に予約されているかチェック
		Call GetAppInfoChassis(conn, rsd, sYMD, sChID)
		If Not rsd.EOF Then
			rsd.Close
			sErrMsg = "指定シャーシはすでに予約されています（" & sChID & "）"
			Exit Function
		End If
		rsd.Close
	End If

	AppUpdCheckDel = True
End Function

'搬入時予約チェック(02/02/09 パラメータADD)
Function AppUpdCheckRec(conn, rsd, _
			sContNoRec, sBKNo, sContSizeRec, sOpeRec, _
			sVslCode, sVoyage, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sContNoRec 		[ I ]搬入コンテナ番号
'	sBKNo			[ I ]搬入ブッキング番号
'	sContSizeRec	[ I ]搬入コンテナサイズ
'	sOpeRec			[ O ]搬入オペレータ
'	sVslCode		[ O ]搬入船名		'02/02/09
'	sVoyage			[ O ]搬入次航		'02/02/09
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheckRec = False

	'ブッキングの有無判定(02/02/09 ADD)
	If Not ChkBooking(conn, rsd, sBKNo, sOpeRec, sVslCode, sVoyage) Then
		sErrMsg = "ブッキングが存在しません（" & sBKNo & "）"
		Exit Function
	End If

	'申請の有無チェック（指定コンテナ）
	If ChkAppInfoExt(conn, rsd, sContNoRec) Then
		sErrMsg = "既に予約済みのコンテナです（" & sContNoRec & "）"
		Exit Function
	End If

	AppUpdCheckRec = True

End Function

'予約枠によるチェック
Function AppUpdCheckTerm(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
			sContNoDel, sChID, sBLNo, sContSizeDel, _
			iAssignCnt, iOpenSlot, sTerm_CD, iDualOpeNo, iOpeOrder, iCnt, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd				[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH				[ I ]予約時間帯（HH）
'	sRDType			[ I ]種類（DUAL,REC,DEL,空白）
'	sContNoRec 		[ I ]搬入コンテナ番号
'	sBKNo			[ I ]搬入ブッキング番号
'	sContSizeRec	[ I ]搬入コンテナサイズ
'	bChkA			[ I ]搬入シャーシに搬出を載せない
'	bChkB			[ I ]搬入シャーシを夕積に使用
'	bChkC  			[ I ]搬入シャーシが20/40兼用シャーシ
'	sContNoDel 		[ I ]搬出コンテナ番号
'	sChID			[ I ]搬出シャーシID
'	sBLNo 			[ I ]搬出ＢＬ番号
'	sContSizeDel	[I/O]搬出コンテナサイズ
'	iAssignCnt 		[ I ]グループ割当数
'	iOpenSlot 		[ I ]開放枠数
'	sTerm_CD		[ I ]搬出入先ターミナルコード
'	iDualOpeNo		[ O ]デュアル作業番号
'	iOpeOrder		[ O ]作業順位
'	iCnt			[ O ]現在往復回数（＝登録グループ枠数）
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheckTerm = False

	'申請情報の取得（指定日、指定時間帯、指定グループ）
	Call GetAppInfoTerm(conn, rsd, sYMD, sHH, sGrpID)
	iCnt = 0			'現在往復回数（＝登録グループ枠数）
	iDualOpeNo = 0		'デュアル作業番号
	iOpeOrder = 0		'作業順位
	Do Until rsd.EOF
		'搬入でデュアルになれる場合(入力されたのが搬入で・該当の搬出相手を捜す)
		If sRDType = "REC" and (Not bChkA) and (Not bChkB) Then
			'ペアになり得る搬出かチェック
				'01/10/19 sta
				'''			If rsd("DualOpeNo") = 0 and _
				'VP対応        		   rsd("RecDel") = "D" and _
				'''			   Trim(rsd("ChassisId")) = "" Then
				'01/10/19 sta
'02/04/27 sta(DUALは同一ターミナルにする)
	   			'If rsd("DualOpeNo") = 0 and rsd("RecDel") <> "R" and _
	   			'   Trim(rsd("ChassisId")) = "" Then
				'	'搬入が香椎・箱崎かまたは、それ以外の前受けの場合は同じVPターミナルの場合
				'	If (sTerm_CD = "KA" OR sTerm_CD = "HK") or _
				'	   (sTerm_CD <> "KA" AND sTerm_CD <> "HK" AND sTerm_CD = Trim(rsd("Terminal")))	  Then	
				'		'兼用かサイズがマッチする場合
				'		If bChkC or _
				'		   (rsd("ContSize") =  "20" and sContSizeRec =  "20") or _
				'		   (rsd("ContSize") <> "20" and sContSizeRec <> "20") Then
				'			iDualOpeNo = rsd("OpeNo")	'デュアル作業番号
				'			iOpeOrder = rsd("OpeOrder") - 1
				'		End If
				'	End If
				'End If

   			If rsd("DualOpeNo") = 0 and rsd("RecDel") <> "R" and _
   			   Trim(rsd("ChassisId")) = "" Then
				'搬入が香椎・箱崎かまたは、それ以外の前受けの場合は同じVPターミナルの場合
				'If (sTerm_CD = "KA" OR sTerm_CD = "HK") or _
				'   (sTerm_CD <> "KA" AND sTerm_CD <> "HK" AND sTerm_CD = Trim(rsd("Terminal")))	  Then	
				'02/04/27 同一ターミナルのDUALにする
				If sTerm_CD = Trim(rsd("Terminal"))	  Then	
					'兼用かサイズがマッチする場合
					If bChkC or _
					   (rsd("ContSize") =  "20" and sContSizeRec =  "20") or _
					   (rsd("ContSize") <> "20" and sContSizeRec <> "20") Then
						iDualOpeNo = rsd("OpeNo")	'デュアル作業番号
						iOpeOrder = rsd("OpeOrder") - 1
					End If
				End If
			End If
		End If
		'搬出でデュアルになれる場合(入力されたのが搬出で・該当の搬入相手を捜す)
		If sRDType = "DEL" and sChID = "" Then	
			'ペアになり得る搬入かチェック
			If rsd("DualOpeNo") = 0 and _
			   rsd("RecDel") = "R" and _
			   rsd("NotDelFlag") <> "Y" and _
			   rsd("NightFlag") <> "Y" Then
				'香椎か箱崎のみ対象となる
'2003/08/25
'''''			If Trim(rsd("Terminal")) = "KA" or Trim(rsd("Terminal")) = "HK" Then
				If Trim(rsd("Terminal")) = "KA" or Trim(rsd("Terminal")) = "IC" Then
'2003/08/25
				'02/04/27 同一ターミナルのDUALにする
					If sTerm_CD = Trim(rsd("Terminal"))	  Then	
						'兼用かサイズがマッチする場合
						If rsd("MixSizeFlag") = "Y" or _
						   (rsd("ContSize") =  "20" and sContSizeDel =  "20") or _
						   (rsd("ContSize") <> "20" and sContSizeDel <> "20") Then
							iDualOpeNo = rsd("OpeNo")	'デュアル作業番号
							iOpeOrder = rsd("OpeOrder") + 1
						End If
					End If
				End If
			End If
		End If
		If rsd("DualOpeNo") = 0 or _
		   rsd("RecDel") <> "R" Then				'=D---> <>R搬入以外に変更(VP対応)
			'搬出は単独のみカウント＝往復回数のカウント
			iCnt = iCnt + 1	'現在往復回数（＝登録グループ枠数）
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	If iDualOpeNo = 0 and _
	   iOpenSlot = 0 and _
	   iCnt >= iAssignCnt Then
		sErrMsg = "割り当て枠がありません"
		Exit Function
	End If

	'他の作業とデュアルにならない場合は作業順位を計算
	If iOpeOrder = 0 Then
		'新規作業順位の取得（指定日、指定時間帯）
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, sRDType)
	End If

	AppUpdCheckTerm = True

End Function
%>
