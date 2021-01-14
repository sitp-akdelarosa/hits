<% 

'空シャーシＳＹ在庫の判定
Function GetVPBooking1(conn, rsd, sVPBook1, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sVPBook1		[ I ]ブッキング番号
'	sErrMsg			[ O ]エラーメッセージ（成功の場合は空白）

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	'SELECT文作成
	sql = "SELECT VslCode, Voyage, LineCode, BookNo, Terminal, ContSize, "
    	sql = sql & " ContType, ContHeight, Material FROM sVPBooking" & sNo
    	sql = sql & " WHERE RTRIM(BookNo) = '" & sVPBook1 & "'" 
	sql = sql & " GROUP BY VslCode, Voyage, LineCode, BookNo, Terminal, "
	sql = sql & " ContSize, ContType, ContHeight, Material ; "
	rsd.Open sql, conn, 0, 1, 1

	IF rsd.EOF Then
		sErrMsg = "該当するブッキング情報がありません"
		rsd.Close
	End If

End Function

'ブッキング予約チェック処理１（作業日のチェック）
Function VPBookCheck1(conn, rsd, _
			sUsrID, sGrpID,	sYMD, sHH, _
			sVPBook, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	sVPBook			[ I ]搬入ブッキング番号
'	sErrMsg			[ O ]エラーメッセージ（成功の場合は空白）

	Dim sWk
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

	End if
End Function

'ブッキング予約更新処理
Function UpdateApp_VPBook(conn, rsd, sUsrID, sGrpID, _
			sYMD, sHH, sVSL, sVOY, sLINE, sVPBook, _
            sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
            sErrMsg, sOpeNoVan)

'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	sVSL			[ I ]船名
'	sVOY			[ I ]次航
'	sLINE			[ I ]航路
'	sVPBook			[ I ]空バンブッキング番号
'	sTERM			[ I ]ターミナルコード
'	sSIZE			[ I ]コンテナサイズ
'	sTYPE			[ I ]タイプ
'	sHIGHT			[ I ]高さ
'	sMATERIAL		[ I ]材質
'	sDeliverTo		[ I ]コンテナ搬出先
'	sErrMsg			[ O ]エラーメッセージ（成功の場合は空白）
'	sOpeNoVan		[ O ]空バン予約番号

	Dim i, sWk, iAssignCnt, bOk, iDualOpeNo, iOpeOrder, iCnt
	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime
	Dim sVslCode, sVoyage, sOpenFlag

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
	End if

	'時間帯指定の場合にはグループの割り当て枠を取得
	iAssignCnt = -1
	sOpenFlag = ""
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'グループの割り当て枠を取得
		iAssignCnt = GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)

		'グループ時間帯情報の取得
		sOpenFlag = GetGrpSlotHH(conn, rsd, sGrpID, sYMD, sHH)

	End If

	'予約チェック
	bOk = AppUpdCheck_VPBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
                             iCnt, sOpeVP, sErrMsg) 

	If bOk Then
		'予約処理
		Call UpdateApp_VPBook2(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
			                 iCnt, sOpeVP, sErrMsg, sOpeNoVan) 
	End If

End Function

'予約チェック
Function AppUpdCheck_VPBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
                             iCnt, sOpeVP, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	sVSL			[ I ]船名
'	sVOY			[ I ]次航
'	sLINE			[ I ]航路
'	sVPBook			[ I ]空バンブッキング番号
'	sTERM			[ I ]ターミナルコード
'	sSIZE			[ I ]コンテナサイズ
'	sTYPE			[ I ]タイプ
'	sHIGHT			[ I ]高さ
'	sMATERIAL		[ I ]材質
'	iAssignCnt 		[ I ]グループ割当数
'	iOpenSlot 		[ I ]開放枠数
'	iDualOpeNo		[ O ]デュアル作業番号
'	iOpeOrder		[ O ]作業順位
'	iCnt			[ O ]予約枠数
'	sOpeVP			[ O ]オペレータ
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheck_VPBook = False

	'抽出可能ブッキングデータチェック
	If Not AppUpdCheckBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, sTERM, _
                             sSIZE, sTYPE, sHIGHT, sMATERIAL, sOpeVP, sErrMsg) Then
		Exit Function
	End If

	'時間帯指定の場合は枠をチェック
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'予約枠によるチェック
		If Not AppUpdCheckTerm_VP(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, sTERM, _
						     sSIZE, sTYPE, sHIGHT, sMATERIAL, iAssignCnt, iOpenSlot, _
						     iDualOpeNo, iOpeOrder, iCnt, sErrMsg) Then

			Exit Function
		End If
	Else
		iDualOpeNo = 0		'デュアル作業番号
		iOpeOrder = 0		'作業順位
		iCnt = 0			'予約枠数
	End If

	AppUpdCheck_VPBook = True
End Function

'抽出可能ブッキングデータチェック
Function AppUpdCheckBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			             sVSL, sVOY, sLINE, sVPBook, sTERM, _
                         sSIZE, sTYPE, sHIGHT, sMATERIAL, sOpeVP, sErrMsg) 
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	sVSL			[ I ]船名
'	sVOY			[ I ]次航
'	sLINE			[ I ]航路
'	sVPBook			[ I ]空バンブッキング番号
'	sTERM			[ I ]ターミナルコード
'	sSIZE			[ I ]コンテナサイズ
'	sTYPE			[ I ]タイプ
'	sHIGHT			[ I ]高さ
'	sMATERIAL		[ I ]材質
'	sOpeVP          	[ O ]オペレータ
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheckBook = False

	Dim sql, sNo
	Dim iRestQtyBook, iCount '予約残数

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	'SELECT文作成
	sql = "SELECT RestQtyBook, Operator FROM sVPBooking" & sNo
    	sql = sql & " WHERE RTRIM(VslCode) = '" & sVSL & "'" 
    	sql = sql & " AND   RTRIM(Voyage) = '" & sVOY & "'" 
    	sql = sql & " AND   RTRIM(LineCode) = '" & sLINE & "'" 
    	sql = sql & " AND   RTRIM(BookNo) = '" & sVPBook & "'" 
    	sql = sql & " AND   RTRIM(Terminal) = '" & sTERM & "'" 
    	sql = sql & " AND   RTRIM(ContSize) = '" & sSIZE & "'" 
    	sql = sql & " AND   RTRIM(ContType) = '" & sTYPE & "'" 
    	sql = sql & " AND   RTRIM(ContHeight) = '" & sHIGHT & "'" 
    	sql = sql & " AND   RTRIM(Material) = '" & sMATERIAL & "'" 
	rsd.Open sql, conn, 0, 1, 1

	If not rsd.EOF Then
		iRestQtyBook = rsd("RestQtyBook")
		sOpeVP = Trim(rsd("Operator"))
	Else
		sErrMsg = "該当するブッキング情報がありません。"
	End If

	rsd.Close

	If sErrMsg <> "" Then
		Exit Function
	End If

	'SELECT文作成(予約本数チェック)
	sql = "SELECT count(OpeNo) AS CountApp FROM sAppliInfo" 
    	sql = sql & " WHERE RTRIM(VslCode) = '" & sVSL & "'" 
    	sql = sql & " AND   RTRIM(Voyage) = '" & sVOY & "'" 
    	sql = sql & " AND   RTRIM(LineCode) = '" & sLINE & "'" 
    	sql = sql & " AND   RTRIM(VPBookNo) = '" & sVPBook & "'" 
    	sql = sql & " AND   RTRIM(Terminal) = '" & sTERM & "'" 
    	sql = sql & " AND   RTRIM(ContSize) = '" & sSIZE & "'" 
    	sql = sql & " AND   RTRIM(ContType) = '" & sTYPE & "'" 
    	sql = sql & " AND   RTRIM(ContHeight) = '" & sHIGHT & "'" 
    	sql = sql & " AND   RTRIM(Material) = '" & sMATERIAL & "'" 
    	sql = sql & " AND   DelFlag <> 'Y' "
    	sql = sql & " AND   Status = '02' "
	rsd.Open sql, conn, 0, 1, 1

	iCount = rsd("CountApp")
	rsd.Close

	If iRestQtyBook <= iCount Then
		sErrMsg = "本数不足により空バンピックできません"
		Exit Function
	End If
	AppUpdCheckBook = True
End Function

'予約枠によるチェック(バンプール)

Function AppUpdCheckTerm_VP(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, sTERM, _
						     sSIZE, sTYPE, sHIGHT, sMATERIAL, iAssignCnt, iOpenSlot, _
						     iDualOpeNo, iOpeOrder, iCnt, sErrMsg)
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	sVSL			[ I ]船名
'	sVOY			[ I ]次航
'	sLINE			[ I ]航路
'	sVPBook			[ I ]空バンブッキング番号
'	sTERM			[ I ]ターミナルコード
'	sSIZE			[ I ]コンテナサイズ
'	sTYPE			[ I ]タイプ
'	sHIGHT			[ I ]高さ
'	sMATERIAL		[ I ]材質
'	iAssignCnt 		[ I ]グループ割当数
'	iOpenSlot 		[ I ]開放枠数
'	iDualOpeNo		[ O ]デュアル作業番号
'	iOpeOrder		[ O ]作業順位
'	iCnt			[ O ]予約枠数
'	sErrMsg			[ O ]エラーメッセージ
'	戻り値			[ O ]OK/NG

	AppUpdCheckTerm_VP = False

	'申請情報の取得（指定日、指定時間帯、指定グループ）
	Call GetAppInfoTerm(conn, rsd, sYMD, sHH, sGrpID)
	iCnt = 0			'現在往復回数（＝登録グループ枠数）
	iDualOpeNo = 0			'デュアル作業番号
	iOpeOrder = 0			'作業順位
	Do Until rsd.EOF
		'搬入でデュアルになれる場合
		'ペアになり得る搬出かチェック

		If rsd("DualOpeNo") = 0 and _
		   rsd("RecDel") = "R" and _
			rsd("NotDelFlag") <> "Y" and _
		   Trim(rsd("ChassisId")) = "" Then
					'搬入が前受けの場合は、対象となる空バン搬出は同じターミナルに限る
					'If (Trim(rsd("Terminal")) = "KA" OR  Trim(rsd("Terminal")) = "HK") OR _
					'   (Trim(rsd("Terminal")) <> "KA" AND Trim(rsd("Terminal")) <> "HK" AND sTERM = Trim(rsd("Terminal"))) Then	
					'	'兼用かサイズがマッチする場合
					'	If (rsd("ContSize") =  "20" and sSIZE =  "20") or _
					'	   (rsd("ContSize") <> "20" and sSIZE <> "20") Then
					'		iDualOpeNo = rsd("OpeNo")	'デュアル作業番号
					'		'01/10/30''
					'		'''iOpeOrder = rsd("OpeOrder") - 1
					'		iOpeOrder = rsd("OpeOrder") + 1
					'	End If
					'End If
			'DUAL同一ターミナルとする(02/04/27)
			If sTERM = Trim(rsd("Terminal")) Then	
				'兼用かサイズがマッチする場合
				If (rsd("ContSize") =  "20" and sSIZE =  "20") or _
				   (rsd("ContSize") <> "20" and sSIZE <> "20") Then
					iDualOpeNo = rsd("OpeNo")	'デュアル作業番号
					'01/10/30''
					'''iOpeOrder = rsd("OpeOrder") - 1
					iOpeOrder = rsd("OpeOrder") + 1
				End If
			End If

		End If

'''		If rsd("DualOpeNo") = 0 or _			'01/10/30 bug
'''		   rsd("RecDel") = "D" Then
		If rsd("DualOpeNo") = 0 or _
		   rsd("RecDel") <> "R" Then
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
		sRDType = "DEL"
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, sRDType)
	End If

	AppUpdCheckTerm_VP = True

End Function


'予約更新処理２
Function UpdateApp_VPBook2(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
			                 iCnt, sOpeVP, sErrMsg, sOpeNoVan) 
'	conn			[ I ]DB接続Object
'	rsd			[ I ]レコードセットObject
'	sUsrID			[ I ]ユーザID
'	sGrpID			[ I ]グループID
'	sYMD			[ I ]予約日（YYYYMMDD）
'	sHH			[ I ]予約時間帯（HH）
'	sVSL			[ I ]船名
'	sVOY			[ I ]次航
'	sLINE			[ I ]航路
'	sVPBook			[ I ]空バンブッキング番号
'	sTERM			[ I ]ターミナルコード
'	sSIZE			[ I ]コンテナサイズ
'	sTYPE			[ I ]タイプ
'	sHIGHT			[ I ]高さ
'	sMATERIAL		[ I ]材質
'	sDeliverTo		[ I ]コンテナ搬出先
'	iAssignCnt 		[ I ]グループ割当数
'	iOpenSlot 		[ I ]開放枠数
'	iDualOpeNo		[ I ]デュアル作業番号
'	iOpeOrder		[ I ]作業順位
'	iCnt			[ I ]予約枠数
'	sOpeVP			[ I ]オペレータ
'	sErrMsg			[ O ]エラーメッセージ
'	sOpeNoVan		[ O ]空バン予約番号

	Dim sql, bTermFlag

	sErrMsg = "ＤＢ更新エラーが発生しました"

	sOpeNoVan = "0"

	'作業番号の発番
	Call GetNewOpeNo(conn, rsd, sYMD, sOpeNoVan)

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
		rsd("DualOpeNo") 	= CLng(sOpeNoVan)		'デュアル作業番号
		rsd("SendFlag")    	= "Y"				'変更送信ﾌﾗｸ
	End If

	rsd.AddNew
	rsd("OpeNo")       = sOpeNoVan				'作業番号
	rsd("UpdtTime")    = Now()				'更新日時
	rsd("UpdtPgCd")    = "SYWB012"				'更新ﾌﾟﾛｸﾞﾗﾑ名
	rsd("UpdtTmnl")    = "WEB"				'更新端末名
	rsd("UserID")      = sUsrID				'ユーザｺｰﾄ
	rsd("GroupID")     = sGrpID				'ｸﾞﾙｰﾌﾟｺｰﾄ
	rsd("AppDate")     = Now()				'申請日
	rsd("WorkDate")    = ChgYMDDate(sYMD)			'作業日
	rsd("RecDel")      = "V"				'作業種類
	rsd("ContNo")      = ""					'ｺﾝﾃﾅ番号
	rsd("BookNo")      = ""					'ﾌﾞｯｷﾝｸﾞ番号
	rsd("Status")      = "02"				'状態
	rsd("Place")       = "VP"				'場所(バンプール)
	rsd("ChassisId")   = ""					'ｼｬｰｼID
	rsd("DelFlag")     = ""					'削除ﾌﾗｸ
	rsd("CReason")    = ""					'キャンセル理由
	rsd("WorkFlag")    = ""					'ｼｬﾄﾙ作業中ﾌﾗｸ
	rsd("SendFlag")    = "Y"				'変更送信ﾌﾗｸ
	If sSIZE = "20" Then
		rsd("Size20Flag")  = "Y"			'20フィートフラグ
	Else
		rsd("Size20Flag")  = ""				'20フィートフラグ
	End If
	rsd("VslCode")     = sVSL				'船名ｺｰﾄ
	rsd("Voyage")      = sVOY				'次航
	rsd("BLNo")        = ""					'ＢＬ番号
	rsd("Term")        = sHH				'時間帯
	If Not bTermFlag Then		'時間帯指定でない場合
		rsd("OpeOrder") = 0				'作業順位
	ElseIf iDualOpeNo > 0 Then	'他の作業とデュアルになる場合
		rsd("OpeOrder") = iOpeOrder			'作業順位
	Else
		rsd("OpeOrder") = iOpeOrder			'作業順位
	End If
	rsd("HopeTerm")    = sHH				'希望時間帯
	If iDualOpeNo > 0 Then	'他の作業とデュアルになる場合
		rsd("DualOpeNo") = iDualOpeNo		'デュアル作業番号
	Else
		rsd("DualOpeNo") = 0				'デュアル作業番号
	End If
	rsd("ContSize")    = sSIZE				'コンテナサイズ
	rsd("NotDelFlag")  = ""					'搬出を載せない
	rsd("NightFlag")   = ""					'夕積シャーシ
	rsd("MixSizeFlag") = ""					'20/40兼用シャーシ
	rsd("DeliverTo")   = sDeliverTo			'搬出先
	rsd("LockFlag")    = ""					'ロックフラグ
	rsd("Operator")    = sOpeVP				'オペレータコード
	rsd("OpeBookNo")   = ""					'自社ブッキング番号
	rsd("ReservNo")    = 0					'予約番号

	rsd("RecContSize") = sSIZE				'ｺﾝﾃﾅｻｲｽﾞ
	rsd("ContHeight")  = sHIGHT				'ｺﾝﾃﾅﾊｲﾄ
	rsd("ContType")    = sTYPE				'ｺﾝﾃﾅﾀｲﾌﾟ
	rsd("Material")    = sMATERIAL			'ｺﾝﾃﾅ材質
	rsd("TareWeight")  = 0					'搬入ｺﾝﾃﾅﾃｱｳｪｲﾄ
	rsd("Gweight")     = 0					'搬入ｺﾝﾃﾅｸﾞﾛｽｳｪｲﾄ
	rsd("SoundUnsound")  = ""				'搬入Sound/Unsound
	rsd("Trucker")  = ""					'搬入運送業者
	rsd("ReceiveFrom")  = ""				'搬入元
	rsd("CustOK")  = ""						'搬入ｺﾝﾃﾅ通関/未通関
	rsd("Terminal")  = sTERM				'該当ターミナル(VP予約)暫定修正
	rsd("VPBookNo")      = sVPBook			'搬出ﾌﾞｯｷﾝｸﾞ番号(VP予約)暫定修正
	rsd("LineCode")      = sLINE			'航路コード(VP予約)暫定修正
	rsd("HeadID")        = ""    			'ヘッドＩＤ(ICCT対応) 03/07/31

	rsd.Update
	rsd.Close
	sErrMsg = ""

End Function

'ターミナル名称の取得
Function GetTerminal_Name(conn, rsd, sTerminal)
'	conn		[ I ]DB接続Object
'	rsd		[ I ]レコードセットObject
'	sTerminal	[ I ]ターミナルコード
'	戻り値		[ O ]名称

	Dim sql
	Dim i

	GetTerminal_Name = ""
	'ターミナル取得
	sql = "SELECT Name FROM sTerminal WHERE RTRIM(Terminal) = '" & sTerminal & "'" & _
		    " AND Umu <> '1'"
	rsd.Open sql, conn, 0, 1, 1

	GetTerminal_Name = Trim(rsd("Name"))

	rsd.Close

End Function

'コンテナタイプ名称の取得
Function GetContType_Name(conn, rsd, sContType)
'	conn		[ I ]DB接続Object
'	rsd		[ I ]レコードセットObject
'	sContType	[ I ]コンテナタイプコード
'	戻り値		[ O ]名称

	Dim sql
	Dim i

	GetContType_Name = ""

	'コンテナタイプ取得
	sql = "SELECT Name FROM sContType WHERE RTRIM(ContType) = '" & sContType & "'"
	rsd.Open sql, conn, 0, 1, 1

	GetContType_Name = Trim(rsd("Name"))
	rsd.Close


End Function

'コンテナ材質取得の取得
Function GetMaterial_Name(conn, rsd, sMaterial)
'	conn		[ I ]DB接続Object
'	rsd		[ I ]レコードセットObject
'	sMaterial	[ I ]コンテナ材質コード
'	戻り値		[ O ]名称

	Dim sql
	Dim i

	'コンテナ材質取得
	sql = "SELECT Name FROM sMaterial WHERE RTRIM(Material) = '" & sMaterial & "'"
	rsd.Open sql, conn, 0, 1, 1

	GetMaterial_Name = Trim(rsd("Name"))
	rsd.Close

End Function

'対象バンプール名称検索
Function SetTerminal(sTerminal, sTerm_Tbl())
	'sTerminal	[ I ]ターミナルコード
	'sTerm_Tbl()	[ I ]ターミナル名称テーブル
	'戻り値		[ O ]対象ターミナル名称

	Dim sStr
	Dim i

	SetTerminal = ""
	for i = 0 to 49
		If sTerminal = Left(sTerm_Tbl(i),3) Then
			SetTerminal = Mid(sTerm_Tbl(i), 4)
			Exit for
		End If
	Next

End Function

'コンテナタイプ名称検索
Function SetContType(sContType, sContType_Tbl())
	'sContType		[ I ]コンテナタイプ
	'sContType_Tbl()	[ I ]コンテナタイプテーブル
	'戻り値			[ O ]対象コンテナタイプ名称

	Dim sStr
	Dim i

	SetContType = ""
	for i = 0 to 14
		If sContType = Left(sContType_Tbl(i),2) Then
			SetContType = Mid(sContType_Tbl(i), 3)
			Exit for
		End If
	Next

End Function

'コンテナ材質名称検索
Function SetMaterial(sMaterial, sMaterial_Tbl())
	'sMaterial	[ I ]コンテナ材質
	'sMaterial()	[ I ]コンテナ材質テーブル
	'戻り値		[ O ]対象コンテナ材質名称

	Dim sStr
	Dim i

	SetMaterial = ""
	for i = 0 to 9
		If sMaterial = Left(sMaterial_Tbl(i),1) Then
			SetMaterial = Mid(sMaterial_Tbl(i), 2)
			Exit for
		End If
	Next

End Function

%>