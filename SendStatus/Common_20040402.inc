<%
'''SQLサーバへの接続
Function ConnDBH(ObjConn,ObjRS)
	Set ObjConn = Server.CreateObject("ADODB.Connection")
	ObjConn.Open "wavedb","wave","wave"
	Set ObjRS = Server.CreateObject("ADODB.Recordset")
	ObjConn.BeginTrans
End Function


'''SQLサーバへの切断
Function DisConnDBH(ObjConn,ObjRS)
	Set ObjRS = Nothing
	ObjConn.CommitTrans
	ObjConn.Close
	Set ObjConn = Nothing
End Function


'''ユーザーＩＤが有効かどうかのチェック
Function CheckLoginH()
	Dim strUserID
	strUserID = Trim(Session.Contents("userid"))

	'''セッションが無効なとき
	If strUserID="" Then
		Dim strtmp
		strtmp = "ログインしてから"&Session.Timeout&"分経過しました。"
		jumpErrorP "0","c000","00","ステータス配信","000",strtmp
	End If
End Function


'''ユーザーＩＤが有効かどうかのチェック
Function CheckLoginH2()
	Dim strUserID
	strUserID = Trim(Session.Contents("userid"))

	'''セッションが無効なとき
	If strUserID="" Then
		Dim strtmp
		strtmp = "ログインしてから"&Session.Timeout&"分経過しました。"
		jumpErrorP "1","c103","01","ステータス配信","000",strtmp
	End If
End Function


'''アクセスログ出力
'''Input :画面ID、画面名、作業ID、確認用オプションデータ
Function WriteLogH(iNum, strScrName, iWrkNum, strOption)
	'サーバ日時の取得
	Dim DayTime,day
	getDayTime DayTime
	'''ユーザタイプコード
	Dim UtypeA
	UtypeA = Array("","荷主","海貨","船社","港運","陸運")
	Dim strDate,logPath,ObjFSO,ObjTS

	strDate = DayTime(0) & DayTime(1) & DayTime(2)
	logPath = Server.MapPath("/log/" & strDate & ".log")

	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set ObjTS = ObjFSO.OpenTextFile(logPath,8,True)
	ObjTS.Write DayTime(0)&"/"&DayTime(1)&"/"&DayTime(2)&" "&DayTime(3)&":"&DayTime(4)&":"&Right("0" & Second(Now), 2)
	ObjTS.Write "," & iNum
	ObjTS.Write "," & strScrName
	ObjTS.Write "," & iWrkNum
	ObjTS.Write "," & UtypeA(Session.Contents("UType")) & "," & Session.Contents("userid")
	ObjTS.Write "," & Request.ServerVariables("SCRIPT_NAME")
	ObjTS.Write "," & Request.ServerVariables("REMOTE_ADDR")
	If strOption<>"" Then
		ObjTS.WriteLine "," & strOption
	Else
		ObjTS.WriteLine 
	End If
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
End Function


'''エラー画面への遷移
'''Input :エラー画面フラグ、画面ID、作業ID、作業名、エラーコード、オプション
Function jumpErrorP(WinFlag,dispId,wkID,wkName,errorCd,etc)
	on error goto 0		'''エラートラップ解除
	Session.Contents("WinFlag")=WinFlag
	Session.Contents("dispId")=dispId
	Session.Contents("wkID")=wkID
	Session.Contents("wkName")=wkName
	Session.Contents("errorCd")=errorCd
	Session.Contents("etc")=etc
	Response.Redirect "./error.asp"
End Function


'''エラー画面への遷移
'''Input :DB接続Object、レコードセットObject、エラー画面フラグ、画面ID、作業ID、作業名、エラーコード、オプション
Function jumpErrorPDB(ObjConn,WinFlag,dispId,wkID,wkName,errorCd,etc)
'''ロールバックとDB切断
	ObjConn.RollbackTrans
	ObjConn.Close
	Set ObjConn = Nothing
	on error goto 0		'エラートラップ解除
	Session.Contents("WinFlag")=WinFlag
	Session.Contents("dispId")=dispId
	Session.Contents("wkID")=wkID
	Session.Contents("wkName")=wkName
	Session.Contents("errorCd")=errorCd
	Session.Contents("etc")=etc
	Response.Redirect "./error.asp"
End Function


'''サーバ日時の取得
'''Input :Array(5)=(YYYY,MM,DD,hh,mm)
Function getDayTime(DayTime)
	Dim MyDate
	MyDate = Now
	DayTime = Array(Year(MyDate),Right("0"&Month(MyDate),2),Right(Date(),2),Right("0"&Hour(MyDate),2),Right("0"&Minute(MyDate),2))
End Function



'''輸入コンテナテーブルのコンテナ存在チェック
'''Input :DB接続Object、レコードセットObject、コンテナ／BL番号、作業フラグ、戻り値
'''Output:ret(存在する=>true　存在しない=>false) 
Function checkImportCont(ObjConn, ObjRS,strNum,Flag,ret)
	Dim StrSQL
	If Flag = 1 Then
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE ContNo=" & strNum
	Else
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE BLNo=" & strNum
	End If
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jumpErrorP "1","b000","00","輸入コンテナテーブル","101","SQL：<BR>"&StrSQL
	end if
	If ObjRS("num") <> 0 Then
		ret = true
	Else
		ret = false
	End If
	ObjRS.close
End Function


'''輸入コンテナテーブルのコンテナ完了チェック
'''Input :DB接続Object、レコードセットObject、コンテナ／BL番号、作業フラグ、戻り値
'''Output:ret(未完了=>true　完了=>false) 
Function checkImportContComp(ObjConn,ObjRS,strNum,Flag,ret)
	dim StrSQL
	If Flag <> 4 Then
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE ContNo=" & strNum 
	Else
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE BLNo=" & strNum 
	End If
	StrSQL = StrSQL & " AND CYDelTime Is Null "
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jumpErrorP "1","b000","00","輸入コンテナテーブル完了チェック","101","SQL：<BR>"&StrSQL
	end if
	If ObjRS("num") <> 0 Then
		ret = true
	Else
		ret = false
	End If
	ObjRS.close
End Function


'''INIファイルからパラメータを読込む
'''Input :Array(1)
'''OutPut:Array(0)=ALARM、Array(1)=COMPLETED_DATA
Function getIni(param)
	Dim ObjFSO,ObjTS,tmpStr
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- ファイルを開く（読み取り専用） ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./INI/PREDEF.INI"),1,false)
	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)
		Select Case tmpStr(0)
			Case "ALARM"
				param(0) = tmpStr(1)
			Case "COMPLETED_DATA"
				param(1) = tmpStr(1)
			Case "INPUT_DATA"
				param(2) = tmpStr(1)
		End Select
	Loop
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
End Function


'''データ受渡しメソッド
'''Input :データ数、データテーブル
Function Set_Data(Num,DtTbl)
	Dim i,j
	Response.Write "  <INPUT type=hidden name='DataNum' value='" & Num & "'>" & vbCrLf
	For i = 0 to Num
		Response.Write "  <INPUT type=hidden name='Datatbl" & i & "' value='"
		For j=0 To UBound(DtTbl(i))-1
			Response.Write DtTbl(i)(j) & ","
		Next
		Response.Write DtTbl(i)(j) & "'>" & vbCrLf
	Next
End Function

'''データ受取りメソッド
'''Input :データ数、データテーブル
Function Get_Data(Num,DtTbl)
	Dim param
	Num=Request("DataNum")
	ReDim DtTbl(Num)
	For Each param In Request.Form
		If Left(param, 7) = "Datatbl" Then
			DtTbl(Mid(param,8))=Split(Request.Form(param), ",", -1, 1)
		End If
	Next
End Function

%>
