<% 
'変更元申請情報の取得
Sub GetOApp(conn, rsd, iOpeNo, sYMD, sErr_msg)
'	conn			[ I ]DB接続Object
'	rsd				[ O ]レコードセットObject
'	iOpeNo			[ I ]作業番号
'	sYMD			[ I ]対象日("YYYYMMDD")
'	sErr_msg		[ O ]エラーメッセージ
	Dim sql

	sErr_msg = ""

	'変更元
	sql = "SELECT * FROM sAppliInfo" & _
			" WHERE OpeNo = " & CStr(iOpeNo) & _
			  " AND WorkDate = '" & cdate(ChgYMDStr(sYMD)) & "'"
	rsd.Open sql, conn, 0, 2, 1		'更新モードでオープン
	If rsd.EOF Then	
		sErr_msg = CStr(iOpeNo) & "　の作業番号は存在しません。"
	End If
	If rsd("Status") = "04" Then
		sErr_msg = CStr(iOpeNo) & "　の作業は完了しています。"
	End If
	If rsd("Status") = "03" Or rsd("DelFlag") = "Y" Then
		sErr_msg = CStr(iOpeNo) & "　の作業は削除されています。"
	End If
	If rsd("Place") = "SY" Then
		sErr_msg = CStr(iOpeNo) & "　の作業はＳＹにあります。"
	End If
	If rsd("WorkFlag") = "Y" Or rsd("LockFlag") = "Y" Then 
		sErr_msg = CStr(iOpeNo) & "　の作業は確定している為変更できません。"
	End If
End Sub

%>