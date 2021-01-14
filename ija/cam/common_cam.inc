<% 
'ＤＢ接続
Function ConnectSvr(conn, rsd) 
	'conn		[ O ]DB接続Object
	'rsd		[ O ]レコードセットObject

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "camdb", "cam", "cam"

	Set rsd = Server.CreateObject("ADODB.Recordset")
End Function

'効果確認用データの出力
Function WriteLog(fs, strScrName, strOption)
	'fs			[ I ]ファイルシステムObject
	'strScrName	[ I ]画面名
	'strOption	[ I ]確認用オプションデータ

    strDate = GetYMDStr(Now)
    logPath = Server.MapPath("log/" & strDate & ".log")
    Set logs=fs.OpenTextFile(logPath, 8, True)
    logs.Write Now
    logs.Write Chr(9) & strScrName
    logs.Write Chr(9) & Session.Contents("userkind") & ":" & Session.Contents("userid")
    logs.Write Chr(9) & Request.ServerVariables("SCRIPT_NAME")
'   logs.Write Chr(9) & Request.ServerVariables("CONTENT_LENGTH")
    logs.Write Chr(9) & Request.ServerVariables("REMOTE_ADDR")
    If strOption<>"" Then
        logs.WriteLine Chr(9) & strOption
    Else
        logs.WriteLine 
    End If

    logs.Close
End Function

'ユーザーＩＤが有効かどうかのチェック
Function CheckLogin(strReturn)
	'strReturn	[ I ]ユーザーＩＤ無効のときの戻り先画面

    ' セッションの有効性をチェック
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))

    ' セッションが無効なとき
    If strUserID="" Then
        ' 戻り画面へリダイレクト
        Response.Redirect "userchk.asp?link=" & strReturn
    End If
End Function

'日付文字列の取得
Function GetYMDStr(DateVal)
	'DateVal	[ I ]日付（Date型）
	'戻り値		[ O ]日付文字列（YYYYMMDD形式）

	GetYMDStr = _
		Year(DateVal) & _
		Right("0" & Month(DateVal), 2) & _
		Right("0" & Day(DateVal), 2)

End Function

'日付文字列の変換
Function ChgYMDStr(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY/MM/DD形式）

	ChgYMDStr = Left(sDate, 4) & "/" & _
				Mid(sDate, 5, 2) & "/" & _
				Mid(sDate, 7, 2)

End Function

'日付文字列の変換２
Function ChgYMDStr2(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY年MM月DD形式）

	ChgYMDStr2 = Left(sDate, 4) & "年" & _
			 	 CStr(CLng(Mid(sDate, 5, 2))) & "月" & _
				 CStr(CLng(Mid(sDate, 7, 2))) & "日"

End Function

'日付文字列の変換３
Function ChgYMDStr3(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY/MM/DD形式）

	Dim sYobi
	ChgYMDStr3 = Left(sDate, 4) & "/" & _
				Mid(sDate, 5, 2) & "/" & _
				Mid(sDate, 7, 2)
	sYobi = sWeekday(Weekday(cDate(ChgYMDStr3)))
	ChgYMDStr3 = ChgYMDStr3 & "(" & sYobi & ")"
End Function

'文字列を日付に変換
Function ChgYMDDate(sYMD)
'	sYMD		[ I ]日付文字列（YYYYMMDD形式）
'	戻り値		[ O ]日付（Date型）

	ChgYMDDate = CDate(ChgYMDStr(sYMD))
End Function

'時間を時間帯に変換
Function GetTimeSlot(iH, iM, sSE)
	'iH			[ I ]時間
	'iM			[ I ]分
	'sSE		[ I ]S:開始、E:終了
	'戻り値		[ O ]時間帯文字列

	Dim sStr

	sStr = CStr(iH) & ":" & Right("0" & CStr(iM), 2)

	If sSE = "S" Then
		sStr = sStr & "～" & CStr(iH + 1) & ":00"
	Else
		sStr = CStr(iH - 1) & ":00～" & sStr
	End If

	GetTimeSlot = sStr

End Function

'曜日を取得する
Function sWeekday(iWk)
'	sYMD		[ I ]数値
'	戻り値		[ O ]曜日

	select case	iWk
        Case 1
            sWeekday =  "日"
        Case 2
            sWeekday =  "月"
        Case 3
            sWeekday =  "火"
        Case 4
            sWeekday =  "水"
        Case 5
            sWeekday =  "木"
        Case 6
            sWeekday =  "金"
        Case 7
            sWeekday =  "土"
	end select
End Function

'固定桁数での表示
Function GetNumStr(vNum, iK)
'	vNum			[ I ]数値（数値or文字列）
'	iK				[ I ]桁数
	GetNumStr = Right(String(iK, "0") & CStr(vNum), iK)
End Function

'日付の表示１
Function DispDateTime(vDate,length)
'	vDate			[ I ]日付
    If VarType(vDate)<>vbDate Then
        strTmp = ""
    Else
        strTmp = _
            Year(vDate) & "/" & _
            Right("0" & Month(vDate), 2) & "/" & _
            Right("0" & Day(vDate), 2) & " " & _
            Right("0" & Hour(vDate), 2) & ":" & _
            Right("0" & Minute(vDate), 2)
        If length=11 Then
            strTmp = Right(strTmp, 11)
        ElseIf length=10 Then
            strTmp = Left(strTmp, 10)
        ElseIf length=8 Then
            strTmp = Mid(strTmp, 3, 8)
        ElseIf length=5 Then
            strTmp = Mid(strTmp, 6, 5)
        End If
        DispDateTime = strTmp
    End If
End Function

'日付の表示２
Function DispDateTimeCell(sCell,length)
'	sCell			[ I ]日付
    If VarType(sCell)<>vbString Or sCell="" Then
        DispDateTimeCell = "<hr width=80%" & ">"
    Else
        If length=5 Then
            strTmp = Mid(sCell, 6, 5)
        ElseIf length=10 Then
            strTmp = Mid(sCell, 6, 5) & "<br>" & Right(sCell, 5)
        ElseIf length=11 Then
            strTmp = Right(sCell, 11)
        Else
            strTmp = sCell
        End If
        DispDateTimeCell = strTmp
    End If
End Function

'予約情報の表示
Function DispReserveCell(sFlag, sCell, sColor)
'	sFlag			[ I ]フラグ
'	sCell			[ I ]日付
'	sColor			[ O ]色
    If sFlag="" Or sFlag="0" Then
        DispReserveCell = "<hr width=80%" & ">"
        sColor=""
    Else
        If sFlag="1" Then
            strTmp = Mid(sCell, 6, 11)
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="2" Then
            strTmp = Mid(sCell, 6, 5) & " 午前"
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="3" Then
            strTmp = Mid(sCell, 6, 5) & " 午後"
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="4" Then
            strTmp = Mid(sCell, 6, 5) & " 夕積"
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="5" Then
            strTmp = "予約削除"
            sColor="<font color='#FF0000'>"
        End If
        DispReserveCell = strTmp
    End If
End Function

'フラグの表示
Function DispFlag(vFlg)
'	vFlg			[ I ]フラグ文字列
    If vFlg=vbNull Or Trim(vFlg)="" Then
        DispFlag = ""
    Else
        If vFlg="1" Then
            DispFlag = "×"
        ElseIf vFlg="2" Then
            DispFlag = "○"
        End If
    End If
End Function

'表の表示
Function DispTableCell(sCell,i)
'	sCell			[ I ]文字列
    If VarType(sCell)<>vbString Or sCell="" Then
        If i=1 Then
            DispTableCell = "<hr width=80%" & ">"
        Else
            DispTableCell = "<br>"
        End If
    Else
        DispTableCell = sCell
    End If
End Function

'Menu選択コンボボックスのJava関数
Sub DispMenuJava
    Response.Write "<!--" & Chr(13) & Chr(10)
    Response.Write "function LinkSelect(form, sel)" & Chr(13) & Chr(10)
    Response.Write "{" & Chr(13) & Chr(10)
    Response.Write "    adrs = sel.options[sel.selectedIndex].value;" & Chr(13) & Chr(10)
    Response.Write "    if (adrs != '-' ) parent.location.href = adrs;" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)

	'コード一覧用function追加	2002/02/06
    Response.Write "function OpenCodeWin()" & Chr(13) & Chr(10)
    Response.Write "{" & Chr(13) & Chr(10)
    Response.Write "	window.open(""codelist.asp"",""codelist"",""scrollbars=yes,resizable=yes,width=250,height=270"");" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)

    Response.Write "// -->" & Chr(13) & Chr(10)
End Sub

'Menu選択コンボボックスの描画
Sub DispMenu
    Response.Write "<FORM action=''>"

	'コード一覧用ボタン追加	2002/02/06
    Response.Write "<input type=button value="" コード一覧 "" onclick=""JavaScript:OpenCodeWin()"">"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

    Response.Write "<SELECT NAME='link' onchange='LinkSelect(this.form, this)'>"
    Response.Write "<OPTION VALUE='#'>Contents"
    Response.Write "<option value='index.asp'>TOP</option>"
    Response.Write "<option value='#'>コンテナ情報照会 </option>"
    Response.Write "<option value='userchk.asp?link=expentry.asp'>├ 輸出コンテナ情報照会 </option>"
    Response.Write "<option value='userchk.asp?link=impentry.asp'>└ 輸入コンテナ情報照会 </option>"
    Response.Write "<option value='#'>各社入力画面</option>"
    Response.Write "<option value='userchk.asp?link=nyuryoku-in1.asp'>├ 船社/ターミナル入力 "
    Response.Write "</option>"
    Response.Write "<option value='userchk.asp?link=nyuryoku-kaika.asp'>├ 海貨入力 </option>"
    Response.Write "<option value='userchk.asp?link=nyuryoku-te.asp'>├ ターミナル入力 </option>"
    Response.Write "<option value='userchk.asp?link=rikuun1.asp'>└ 陸運入力</option>"
    Response.Write "<option value='userchk.asp?link=sokuji.asp'> 即時搬出システム </option>"
    Response.Write "<option value='userchk.asp?link=hits.asp'>ストックヤード利用システム</option>"
    Response.Write "<option value='userchk.asp?link=terminal.asp'>ゲート前映像・混雑状況照会 </option>"
    Response.Write "<option value='userchk.asp?link=request.asp'>利用者アンケート・Ｑ＆Ａ</option>"
    Response.Write "</SELECT>"
    Response.Write "</FORM>"
End Sub

'Menuバーの描画
Sub DispMenuBar
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td valign='bottom' align='right'><a href='index.asp'><img src='gif/b-home.gif' border='0' width='270' height='23' usemap='#map'></a></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td bgcolor='000099' height='10'><img src='gif/1.gif'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

'戻り先の登録
Sub DispMenuBarBack(strBackURL)
'    ' 戻り先を登録
'    Dim strBackURL
'    strBackURL = Trim(Session.Contents("returnurl"))
'    If strBackURL="" Then
'        strBackURL = "index.asp"
'    End If

    Response.Write "<map name='map'>"
    Response.Write "<area shape='poly' coords='20,0,152,0,134,22,0,22' href='" & strBackURL & "'>"
    Response.Write "<area shape='poly' coords='154,0,136,22,284,22,284,0' href='index.asp'>"
    Response.Write "</map>"

'    Session.Contents("returnurl") = Request.ServerVariables("SCRIPT_NAME")
End Sub

'エラーメッセージの表示
Sub DispErrorMessage(strMessage)
    Response.Write "<center>"
    Response.Write "<table border=1 cellpadding='2' cellspacing='1'>"
    Response.Write "<tr>"
    Response.Write "<td bgcolor='#FFFFFF'>"
    Response.Write "<table border='0'>"
    Response.Write "<tr>"
    Response.Write "<td valign='middle'><img src='gif/error.gif'></td>"
    Response.Write "<td><b><font color='#FF4400'>"
    Response.Write strMessage
    Response.Write "</font></b></td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</center>"
End Sub

'注目メッセージの表示
Sub DispInformationMessage(strMessage)
    Response.Write "<center>"
    Response.Write "<table border=1 cellpadding='2' cellspacing='1'>"
    Response.Write "<tr>"
    Response.Write "<td bgcolor='#FFFFFF'>"
    Response.Write "<table border='0'>"
    Response.Write "<tr>"
    Response.Write "<td valign='middle'><img src='gif/info.gif'></td>"
    Response.Write "<td><b><font color='#0000FF'>"
    Response.Write strMessage
    Response.Write "</font></b></td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</center>"
End Sub

' トランザクション　日毎シーケンス番号発行
Function GetDailyTransNo()
	'戻り値		[ O ]日毎シーケンス番号(文字列)

	' 桁数
	Const FIGURE_NUM = 6

	If Application("LastTransDate") <> Date Or Application("DailyTransNo") = "" Then
		' システム起動後または日付が変わった場合は１を設定
		Application("LastTransDate") = Date
		Application("DailyTransNo") = 1
	Else
		' 通常は１ずつ加算
		Application("DailyTransNo") = Application("DailyTransNo") + 1
	End If 

	' 決められた桁数の文字列に
	Dim sNo

	sNo = CStr(Application("DailyTransNo"))
	If FIGURE_NUM < Len(sNo) Then
		sNo = Right(sNo, FIGURE_NUM)
	Else
		Do While Len(sNo) < FIGURE_NUM 
			sNo = "0" & sNo
		Loop				
	End If

	GetDailyTransNo = sNo
End Function

'Tempファイルの種類を設定
Sub SetTempFile(strKind)
    Session.Contents("tempfilekind")=strKind
End Sub

'Tempファイルが有効かどうかのチェック
Function CheckTempFile(strKind,strReturn)
	'strKind	[ I ]Tempファイルの種類
	'strReturn	[ I ]Tempファイル無効のときの戻り先画面

    ' Tempファイルの有効性をチェック
    Dim strTempKind
    strTempKind = Trim(Session.Contents("tempfilekind"))

    ' Tempファイルが無効なとき
    If strTempKind<>strKind Then
        ' 戻り画面へリダイレクト
        Response.Redirect strReturn
    End If
End Function

'日時２桁変換
Function DateFormat(strNum)
	If strNum="1" Then
		strNum = "01"
	End If
	If strNum="2" Then
		strNum = "02"
	End If
	If strNum="3" Then
		strNum = "03"
	End If
	If strNum="4" Then
		strNum = "04"
	End If
	If strNum="5" Then
		strNum = "05"
	End If
	If strNum="6" Then
		strNum = "06"
	End If
	If strNum="7" Then
		strNum = "07"
	End If
	If strNum="8" Then
		strNum = "08"
	End If
	If strNum="9" Then
		strNum = "09"
	End If
    DateFormat = strNum
End Function

'/*
' *	場所表示用の文字列を返す
' *
' *	@param	String	strScriptName	ASPファイル名
' *	@return	String	strRoute		Topから現在のページまでの道筋
' */
Function SetRoute(strScriptName)
	Dim bRouteFlag,iIndex,strParent,strSubName,strRoute,strExpCont,strImpCont,strTerminal

	iIndex = InStrRev(strScriptName,"/")
	strScriptName = Right(strScriptName,Len(strScriptName)-iIndex)

	strExpCont 	= "輸出コンテナ情報照会（作業選択） "
	strImpCont 	= "輸入コンテナ情報照会（作業選択） "
	strTerminal = "ターミナル所用時間照会 "
	strKaika 	= "海貨入力（作業選択） "
	strVslTerm 	= "船社／ターミナル入力（作業選択） "
	strTerm 	= "ゲート前映像・混雑状況照会 "
	strTrack 	= "陸運入力（コンテナNo.入力） "
	strGate 	= "ゲート通行時間予約システム "
	strReq	 	= "利用者アンケート "

	bRouteFlag = true
	strParent  = ""

    Select Case strScriptName

	'輸出コンテナ情報照会
        Case "expentry.asp","expcntnr.asp","expreload.asp"
								strSubName = strExpCont
								bRouteFlag = false
        Case "expcsv.asp","expcsvin.asp"
								strSubName = "輸出コンテナ情報照会用ファイル転送画面 "
								strParent  = strExpCont
		Case "expdetail.asp"	strSubName = "輸出コンテナ照会（単独コンテナ） "
        Case "explist.asp"		strSubName = "輸出コンテナ照会（複数コンテナ） "
								strParent  = strExpCont
        Case "explist1.asp"		strSubName = "位置情報 "
								strParent  = "輸出コンテナ照会（複数コンテナ） "
        Case "explist2.asp"		strSubName = "コンテナ作成に係る情報 "
								strParent  = "輸出コンテナ照会（複数コンテナ） "
        Case "explist3.asp"		strSubName = "ターミナル、本船に係る情報 "
								strParent  = "輸出コンテナ照会（複数コンテナ） "

		'輸出コンテナ情報照会（松下仕様）
        Case "ms-expentry.asp"
								strKind = Request.QueryString("kind")
								Select Case strKind
									Case 1	strSubName = "海貨照会 "
									Case 2	strSubName = "陸運照会 "
									Case 3	strSubName = "荷主照会 "
								End Select
								strParent  = strExpCont
        Case "ms-explist1.asp","ms-explist2.asp","ms-explist3.asp"
								strSubName = "輸出コンテナ情報一覧 "
        Case "ms-expdetail.asp"	strSubName = "コンテナ情報詳細 "
								strParent  = "輸出コンテナ情報一覧 "
        Case "ms-expinput.asp","ms-expinput-syori.asp"
								strSubName = "時刻入力 "

	'輸入コンテナ情報照会
        Case "impentry.asp","impcntnr.asp","impreload.asp"
								strSubName = strImpCont
								bRouteFlag = false
        Case "impcsv.asp","impcsvin.asp"
								strSubName = "輸入コンテナ情報照会用ファイル転送画面 "
								strParent  = strImpCont
		Case "impdetail.asp"	strSubName = "輸入コンテナ照会（単独コンテナ） "
        Case "implist.asp"		strSubName = "輸入コンテナ照会（複数コンテナ） "
								strParent  = strImpCont
        Case "implist1.asp"		strSubName = "ターミナル搬入までの位置情報 "
								strParent  = "輸入コンテナ照会（複数コンテナ） "
        Case "implist2.asp"		strSubName = "ターミナル搬出後の基本情報および位置情報 "
								strParent  = "輸入コンテナ照会（複数コンテナ） "

		'輸入コンテナ情報照会（松下仕様）
        Case "ms-impentry.asp"
								strKind = Request.QueryString("kind")
								Select Case strKind
									Case 1	strSubName = "海貨照会 "
									Case 2	strSubName = "陸運照会 "
									Case 3	strSubName = "荷主照会 "
								End Select
								strParent  = strImpCont
        Case "ms-implist1.asp","ms-implist2.asp"
								strSubName = "輸入コンテナ情報一覧 "
        Case "ms-impdetail.asp"	strSubName = "コンテナ情報詳細 "
								strParent  = "輸入コンテナ情報一覧 "
        Case "ms-impinput.asp","ms-impinput-syori.asp"
								strSubName = "時刻入力 "

	'ターミナル所用時間照会
	'ターミナル所用時間照会
        Case "terminal.asp"		strSubName = strTerm
								bRouteFlag = false
        Case "camera.asp"		strSubName = "ゲート周辺映像 "
								strParent  = strTerm
        Case "photogate.asp"	strSubName = "ゲート前映像 "
								strParent  = "ゲート周辺映像 "
        Case "phototaikijyo.asp"	strSubName = "待機場映像 "
									strParent  = "ゲート周辺映像 "
        Case "photokamome.asp"	strSubName = "かもめ大橋映像 "
								strParent  = "ゲート周辺映像 "

	'共通ログイン画面
        Case "userchk.asp"
							If InStr(Request.QueryString("link"),"-impentry.asp")<>0 Then
								strSubName = strImpCont & "&gt; ログイン"
								Session.Contents("route") = "Top &gt; " & strImpCont
							ElseIf InStr(Request.QueryString("link"),"-expentry.asp")<>0 Then
								strSubName = strExpCont & "&gt; ログイン"
								Session.Contents("route") = "Top &gt; " & strExpCont
							Else
								strSubName = "ログイン"
								Session.Contents("route") = ""
							End If
							bRouteFlag = false

	'船社／ターミナル入力
        Case "nyuryoku-in1.asp"				strSubName = strVslTerm
											bRouteFlag = false
        Case "nyuryoku-csv.asp","nyuryoku-csvin.asp"
											strSubName = "船社／ターミナル用ファイル転送画面 "
        Case "nyuryoku-port.asp"			strSubName = "本船動静一覧 "
        Case "nyuryoku-vsl.asp"				strSubName = ""
        Case "nyuryoku-sch.asp"				strSubName = "本船動静入力（更新）"
        Case "nyuryoku-new.asp"				strSubName = "本船動静入力（新規）"
        Case "nyuryoku-new-ist.asp"			strSubName = ""

	'海貨入力
        Case "nyuryoku-kaika.asp"			strSubName = strKaika
											bRouteFlag = false
		'ＣＹ搬入日指示
        Case "nyuryoku-ex.asp"				strSubName = "輸出ＣＹ搬入日指示 "
        Case "nyuryoku-ex-csv.asp","nyuryoku-ex-csvin.asp"
											strSubName = "輸出ＣＹ搬入日指示用ファイル転送画面 "
        Case "nyuryoku-ex-syori.asp"		strSubName = ""

		'シール番号・重量
        Case "nyuryoku-ki.asp"				strSubName = "輸出シール番号・重量 "
        Case "nyuryoku-kcsv.asp","nyuryoku-kcsvin.asp"
											strSubName = "輸出シール番号・重量用ファイル転送画面 "
        Case "nyuryoku-ki-syori.asp"		strSubName = ""

		'実入り倉庫届時刻指示
        Case "nyuryoku-im.asp"				strSubName = "輸入実入り倉庫届け時刻指示 "
        Case "nyuryoku-im-csv.asp","nyuryoku-im-csvin.asp"
											strSubName = "輸入実入り倉庫届け時刻指示用ファイル転送画面 "
        Case "nyuryoku-im-syori.asp"		strSubName = ""

		'輸出貨物情報
        Case "ms-kaika-expinfo.asp"			strSubName = "輸出貨物情報入力 "
        Case "ms-kaika-expinfo-new.asp"
											If Request.QueryString("kind")<>0 Then
												strSubName = "新規入力 "
											Else
												strSubName = "更新入力 "
											End If
        Case "ms-kaika-expinfo-list.asp"	strSubName = "更新対象一覧 "
        Case "ms-kaika-expinfo-csv.asp","ms-kaika-expinfo-csvin.asp"
											strSubName = "貨物情報入力用ファイル転送画面 "
        Case "ms-kaika-expinfo-exec.asp","ms-kaika-expinfo-updatecheck.asp"
											strSubName = ""

		'輸出コンテナ情報
        Case "ms-kaika-expcontinfo.asp"		strSubName = "輸出コンテナ情報入力 "
        Case "ms-kaika-expcontinfo-new.asp"
											If Request.QueryString("kind")<>0 Then
												strSubName = "新規入力 "
											Else
												strSubName = "更新入力 "
											End If
        Case "ms-kaika-expcontinfo-list.asp"	strSubName = "更新対象一覧 "
        Case "ms-kaika-expcontinfo-csv.asp","ms-kaika-expcontinfo-csvin.asp"
												strSubName = "コンテナ情報入力用ファイル転送画面 "
        Case "ms-kaika-expcontinfo-exec.asp","ms-kaika-expcontinfo-updatecheck.asp"
												strSubName = ""

		'輸入コンテナ情報
        Case "ms-kaika-impcontinfo.asp"		strSubName = "輸入コンテナ情報入力 "
        Case "ms-kaika-impcontinfo-new.asp"
											If Request.QueryString("kind")<>0 Then
												strSubName = "新規入力 "
											Else
												strSubName = "更新入力 "
											End If
        Case "ms-kaika-impcontinfo-list.asp"	strSubName = "更新対象一覧 "
        Case "ms-kaika-impcontinfo-csv.asp","ms-kaika-impcontinfo-csvin.asp"
												strSubName = "コンテナ情報入力用ファイル転送画面 "
        Case "ms-kaika-impcontinfo-exec.asp","ms-kaika-impcontinfo-updatecheck.asp"
												strSubName = ""

	'ターミナル入力
		Case "nyuryoku-te.asp"				strSubName = strTerm
											bRouteFlag = false
		Case "nyuryoku-te1.asp"
											If strChoice="bl" Then
												strSubName = "搬入確認予定時刻入力( BL単位 ) "
											Else
												strSubName = "搬入確認予定時刻入力( 本船単位 ) "
											End If
		Case "nyuryoku-tmnl-csv.asp","nyuryoku-tmnl-csvin.asp"
											strSubName = "搬入確認予定時刻入力用ファイル転送画面 "

	'陸運入力
		Case "rikuun1.asp"					strSubName = strTrack
											bRouteFlag = false
		Case "rikuun2.asp","rikuun3.asp"	strSubName = "完了作業送信 "
											strParent  = strTrack

	'ゲート通行時間予約システム
		Case "gate.asp"						strSubName = strGate
											bRouteFlag = false

	'利用者アンケート
		Case "request.asp"					strSubName = strReq
											bRouteFlag = false

		Case Else
	End Select


	If bRouteFlag Then

		If strSubName="" Then
			strRoute = Session.Contents("route")
		ElseIf strParent<>"" Then
			strSubName = " &gt; " & strSubName
			strRoute = Session.Contents("route")
			iIndex = InStr(strRoute,strParent)
			strRoute = Left(strRoute,iIndex-1) & strParent & strSubName
		Else
			strSubName = " &gt; " & strSubName
			strRoute = Session.Contents("route")
			iIndex = InStr(strRoute,strSubName)
			If iIndex<>0 Then
				strRoute = Left(strRoute,iIndex) & strSubName
			Else
				strRoute = strRoute &strSubName
			End If
		End If

	Else
		strRoute = "Top &gt; " & strSubName
	End If

	SetRoute = strRoute

End Function

Function GetUpdateTime(fs)

	Dim strPath,f
    strFileName = Session.Contents("tempfile")

	If strFileName<>"" Then
	    strFileName="./temp/" & strFileName
		strPath = Server.MapPath(strFileName)
	Else
		GetUpdateTime = ""
		Exit Function
	End If

	Set f = fs.GetFile(strPath)
	dateTimeTmp = f.DateLastModified
	GetUpdateTime = Year(dateTimeTmp) & "年" & _
		Right("0" & Month(dateTimeTmp), 2) & "月" & _
		Right("0" & Day(dateTimeTmp), 2) & "日" & _
		Right("0" & Hour(dateTimeTmp), 2) & "時" & _
		Right("0" & Minute(dateTimeTmp), 2) & "分現在の情報"

End Function

%>