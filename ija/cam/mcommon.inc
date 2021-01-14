<%
' ユーザIDの形式チェック
Function CheckUserID(sUserID)
	'sUserID		[ I ]ユーザID
	'CheckUserID	[ O ]エラーメッセージ

	' 運送会社会社コード(ユーザIDの先頭)の最大値
	Const COMPANY_MAX = "008"

	Dim sCompany, sDriver

	If sUserID = "" Then
		CheckUserID = "ユーザーID未入力"
		Exit Function
	End If

	If (Not IsNumeric(sUserID)) Or (Len(sUserID) <> 6) Then
		CheckUserID = "不正なユーザーID<br>[　" & sUserID & "　]"
		Exit Function
	End If

	sCompany = Left(sUserID, Len(COMPANY_MAX))
	sDriver = Right(sUserID, Len(sUserID) - Len(COMPANY_MAX))
	If CInt(sCompany) = 0 Or CInt(sDriver) = 0 Or CInt(COMPANY_MAX) < CInt(sCompany) Then
		CheckUserID = "不正なユーザーID<br>[　" & sUserID & "　]"
		Exit Function
	End If

	CheckUserID = ""
End function

' 携帯電話の種別を取得
Function GetPhoneType()
	'GetPhoneType		[ O ]携帯電話の種別
	'				I →i-mode
	'				J →J-Sky
	'				E →EzWeb
	'				P →PC
	'				空→不明

	Dim str			' ユーザエージェント
	
	' ユーザエージェントを取得
	str = Request.ServerVariables("HTTP_USER_AGENT")
	
	' ユーザエージェント毎に処理
	If Instr(1, str, "DoCoMo", 1) > 0 Then
		' ユーザエージェントがi-modeの場合
		GetPhoneType = "I"
		Exit Function
	End If

	If Instr(1, str, "J-PHONE", 1) > 0 Then
		' ユーザエージェントがJ-SKYの場合
		GetPhoneType = "J"
		Exit Function
	End If

    '2010/03/18 Add-S NICS okimi	携帯の追加
	If Instr(1, str, "Vodafone", 1) > 0 Then
		' ユーザエージェントがVodafoneの場合
		GetPhoneType = "J"
		Exit Function
	End If

	If Instr(1, str, "SoftBank", 1) > 0 Then
		' ユーザエージェントがソフトバンクの場合
		GetPhoneType = "J"
		Exit Function
	End If
    '2010/03/18 Add-E NICS okimi

	If Instr(1, str, "Mozilla", 1) > 0 Then
		' ユーザエージェントがPCの場合
		GetPhoneType = "P"
		Exit Function
	End If

	If Instr(1, str, "UP.Browser", 1) > 0 Then
		' ユーザエージェントがEZ-WEBの場合
		GetPhoneType = "E"
		Exit Function
	End If

	GetPhoneType = ""
End Function

' 携帯電話の種別に応じたTitleタグの内容を取得
Function GetTitleTag(strTitle)
	'stTitle		[ I ]タイトル
	'GetTitleTag	[ O ]タイトルタグ
	
	' J-Skyはタイトルタグが画面に表示されるので指定しない
	If GetPhoneType() = "J" Then
		GetTitleTag = ""
	Else
		GetTitleTag = "<title>" & strTitle & "</title>"
	End If
End Function

' 携帯電話の種別に応じた画像の拡張子を取得
Function GetImageExt()
	'GetImageExt	[ O ]画像の拡張子(.付)
	
	Select Case GetPhoneType()
		Case "I":
			GetImageExt = ".gif"
		Case "J":
			GetImageExt = ".png"
		Case "E":
			GetImageExt = ".bmp"
		Case "P":
			GetImageExt = ".gif"
		Case Else:
			GetImageExt = ".gif"
	End Select
End Function

' 携帯電話の種別に応じた数字ボタンリンク用のタグを取得
Function GetKeyTag(strKeyName)
	'strKeyName		[ I ]ボタン名
	'GetKeyTag		[ O ]数字ボタンリンク用のタグ
	
	Select Case GetPhoneType()
		Case "I":
			GetKeyTag = "accesskey=""" & Right(strKeyName,1) & """"
		Case "J":
			GetKeyTag = "directkey=""" & Right(strKeyName,1) & """"
		Case "E":
			GetKeyTag = "accesskey=""" & Right(strKeyName,1) & """"
		Case Else:
			GetKeyTag = ""
	End Select
End Function

' 携帯電話の種別に応じた数字ボタンのラベルを取得
Function GetKeyLabel(strKeyName)
	'strKeyName		[ I ]ボタン名
	'GetKeyLabel	[ O ]数字ボタンリンク用のラベル
	
	Select Case GetPhoneType()
		Case "I":
'	NECの機種が対応していないので絵文字の使用を中止 2002/02/14
'			If IsNumeric(strKeyName) Then
'				GetKeyLabel = "&#" & CStr(63878 + CInt(strKeyName))
'			Else
				GetKeyLabel = Right(strKeyName,1) & "."
'			End If
		Case "P":
			GetKeyLabel = "[" & Right(strKeyName,1) & "]"
		Case Else:
			GetKeyLabel = ""
	End Select
End Function

' 携帯電話の種別に応じたテキストボックスのサイズ・文字種のタグを取得
Function GetTextSizeMode(nSize, strMode)
	'nSize				[ I ]サイズ
	'strMode			[ I ]文字種(H:かな, K:カナ, A:英字, N:数字)
	'GetTextSizeMode	[ O ]サイズ・文字種のタグ
	
	Dim sModeName
	Select Case GetPhoneType()
		Case "I":
			Select Case strMode
				Case "H":
					 sModeName = "1"
				Case "K":
					 sModeName = "2"
				Case "A":
					 sModeName = "3"
				Case "N":
					 sModeName = "4"
			End Select
			GetTextSizeMode = " size=""" & CStr(nSize) & """ istyle=""" & sModeName & """ "  
		Case "J":
			Select Case strMode
				Case "H":
					 sModeName = "hiragana"
				Case "K":
					 sModeName = "katakana"
				Case "A":
					 sModeName = "alphabet"
				Case "N":
					 sModeName = "numeric"
			End Select
			GetTextSizeMode = " size=""" & CStr(nSize) & """ mode=""" & sModeName & """ "
		Case Else:
			GetTextSizeMode = " size=""" & CStr(Int(nSize * 1.1 + 3)) & """ " 
	End Select
End Function

' 搬出可能か(輸入コンテナテーブルから判断)
Function CanCarryOut(rs) 
	'rs				[ I ]輸入コンテナテーブルのレコ－ドセット
	'CanCarryOut	[ O ]搬出可能か('Y':可,'S':済,' ':不可)

	Dim tNow
	tNow = Now

    '日付までしか入っていないDateTime型の比較準備 20040224 S
    Dim strchkNow, strchkFreeTime, strchkDelPermitDate
    Dim strchkOLTDateFrom, strchkOLTDateTo, strchkFreeTimeExt
    strchkNow = DispDateTime( Now, 8 )
    strchkFreeTime = DispDateTime( rs("FreeTime"), 8 )
    strchkDelPermitDate = DispDateTime( rs("DelPermitDate"), 8 )
    '搬出可能日がNULLのとき搬出可能になるのを防ぐ(2004/09/01 Add)
    If strchkDelPermitDate = "" Then
        strchkDelPermitDate = "99999999"
    End If
    '(2004/09/01 Add End)
    strchkOLTDateFrom = DispDateTime( rs("OLTDateFrom"), 8 )
    strchkOLTDateTo = DispDateTime( rs("OLTDateTo"), 8 )
    strchkFreeTimeExt = DispDateTime( rs("FreeTimeExt"), 8 )
    '日付までしか入っていないDateTime型の比較準備 20040224 E

    '--- add by mes(2005/3/28)
    ' オンドックで卸リスト対象外フラグが対象外でないなら卸リスト番号の有無をチェック
    If Trim(rs("OffDockFlag"))="N" And (Trim(rs("DsListFlg"))<>"1" Or IsNull(Trim(rs("DsListFlg"))) = True ) Then
    	If Trim(rs("ListNo"))="" Then
    		CanCarryOut = " "
			Exit Function
    	End If
    End If
    '--- end mes

	' 通関／保税輸送のチェック
	If Trim(rs("OLTICFlag")) = "I" Then
		If Trim(rs("OLTICNo")) = "" Then	
'			CanCarryOut = False
			CanCarryOut = " "
			Exit Function
		End If
	'OLTICFlagが空白のとき、日付チェックを行わず、通関OKとする(2005/01/18 Chenge)
'   Else
	ElseIf Trim(rs("OLTICFlag"))<>"" Then
	'OLTICFlagが空白のとき、日付チェックを行わず、通関OKとする(2005/01/18)
'		If Not (tNow >= rs("OLTDateFrom") And tNow <= rs("OLTDateTo")) Then	'D20040224
		If Not (strchkNow >= strchkOLTDateFrom And strchkNow <= strchkOLTDateTo) Then 'I20040224
'			CanCarryOut = False
			CanCarryOut = " "
			Exit Function
		End If		
	'OLTICFlagが空白のとき、許可日と許可番号のチェックをし、通関OKとする(2005/02/03 mod)
    Else
		If DispDateTime( rs("OLTICDate"), 8 ) <> "" And Trim(rs("OLTICNo")) = "" Then
			CanCarryOut = "Y"
		End If
	'OLTICFlagが空白のとき、日付チェックを行わず、通関OKとする(2005/01/18)
	End If

'	If Not IsNull(rs("CYDelTime")) Then
	If DispDateTime(rs("CYDelTime"),0)<>"" Then
		CanCarryOut = "S"
'	ElseIf Trim(rs("DOStatus")) = "Y" And _
'		(rs("FreeTime") >= tNow Or rs("FreeTimeExt") >= tNow) And _
'		rs("DelPermitDate") <= tNow Then	'D20040224
'--- mod by MES(2004/9/10)
'	ElseIf Trim(rs("DOStatus")) = "Y" And _
'		(strchkFreeTime >= strchkNow Or strchkFreeTimeExt >= strchkNow) And _
'		strchkDelPermitDate <= strchkNow Then	'I20040224
	ElseIf Trim(rs("DOStatus")) = "Y" And _
		(strchkFreeTime >= strchkNow Or strchkFreeTimeExt >= strchkNow) And _
		strchkDelPermitDate <= strchkNow AND _
		Trim(rs("OLTICNo")) <> "" AND DispDateTime( rs("OLTICDate"), 8 ) <> "" Then
'--- end MES
		
		CanCarryOut = "Y"
	Else

		CanCarryOut = " "
	End If
End Function

' 数値を指定した桁数の文字列に変換(右詰・余白には０)
Function ArrangeNum(nNumber, nFigure)
	'nNumber	[ I ]変換する数値
	'nFigure	[ I ]変換後の桁数
	'ArrangeNum	[ O ]変換後の文字列

	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNum = sNum
End Function

'効果確認用データの出力(携帯電話用)
Function WriteLogM(fs, strUserID, iNum, strScrName, iWrkNum, strPhoneKind,strOption)
	'fs			[ I ]ファイルシステムObject
	'strUserID	[ I ]ユーザID
	'iNum		[ I ]画面ID
	'strScrName	[ I ]画面名
	'iWrkNum	[ I ]作業ID
	'strPhoneKind [ I ]携帯機種
	'strOption	[ I ]確認用オプションデータ


	' Logを出力するフォルダ
	Const LOG_FOLDER = "../log"

	' Logファイルの拡張子 
	Const LOG_EXTENT = "log"

	Dim strDate,strNowTime
	Dim logPath
	Dim logs

	strNowTime = Year(Now) & "/" & _
		Right("0" & Month(Now), 2) & "/" & _
		Right("0" & Day(Now), 2) & " " & _
		Right("0" & Hour(Now), 2) & ":" & _
		Right("0" & Minute(Now), 2) & ":" & _
		Right("0" & Second(Now), 2)
	strDate = GetYMDStr(Now)
	logPath = Server.MapPath(LOG_FOLDER & "/" & strDate & "." & LOG_EXTENT)
	Set logs=fs.OpenTextFile(logPath, 8, True)
	logs.Write strNowTime
    logs.Write "," & iNum
	logs.Write "," & strScrName
    logs.Write "," & iWrkNum
	If strPhoneKind="E" Then
		logs.Write ",A"
	Else
		logs.Write "," & strPhoneKind
	End If
	logs.Write "," & strUserID
	logs.Write "," & Request.ServerVariables("SCRIPT_NAME")
'	logs.Write "," & Request.ServerVariables("CONTENT_LENGTH")
	logs.Write "," & Request.ServerVariables("REMOTE_ADDR")
	If strOption<>"" Then
		logs.WriteLine "," & strOption
	Else
		logs.WriteLine 
	End If

	logs.Close

    '2009/03/05 Add-S Tanaka
     xLogReg iNum,iWrkNum,strUserID
    '2009/03/05 Add-E Tanaka

End Function

'2009/03/05 Add-S Tanaka
'ログをDBに登録
Function xLogReg(iNum,iWrkNum,strUserId)
	'iNum		[ I ]画面ID
	'iWrkNum	[ I ]作業ID

	Dim ObjConn, ObjRS, StrSQL,fs,ti,strTemp,anyTmpTitle
	Dim iAddFlag
	Dim PageNum(),WkNum(),PageTitle(),SubTitle()
	Dim strTitleFileName



	'DB接続
	ConnectSvr ObjConn,ObjRS


	Set fs=Server.CreateObject("Scripting.FileSystemObject")

	' ログタイトル取得
	strTitleFileName="/logija.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iAddFlag=0


	'ログタイトル分ループ
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		'ログタイトルに該当するもの場合
		IF TRIM(anyTmpTitle(0))=TRIM(CSTR(iNum)) and TRIM(anyTmpTitle(1))=TRIM(CSTR(iWrkNum)) then
			'ログインコードが該当のもの
			IF strUserId<>"90000" and strUserId<>"10000" and strUserId<>"MES" and strUserId<>"MEO" Then
				iAddFlag=1
				Exit Do
			ELSE
				Exit Function
			End If
		End If
	Loop
	ti.Close

	'登録可能なデータが存在する場合
	If iAddFlag=1 Then

		StrSQL = ""
		StrSQL = StrSQL & " SELECT * FROM xLog "
		StrSQL = StrSQL & " WHERE LogDate='" & GetYMDStr(Now) &"'"
		StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
		StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		
		ObjRS.Open StrSQL, ObjConn

		'レコードが存在する場合
		If ObjRS.EOF=False then
			StrSQL = ""
			StrSQL = StrSQL & " UPDATE xLog SET "
			StrSQL = StrSQL & " DataCount=DataCount+1 "
			StrSQL = StrSQL & " ,UpdtTime=GetDate() "
			StrSQL = StrSQL & " ,UpdtPgCd='MCOMNINC' "
			StrSQL = StrSQL & " ,UpdtTmnl='" & Session.Contents("userid") & "'"
			StrSQL = StrSQL & " WHERE LogDate='" & GetYMDStr(Now) &"'"
			StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
			StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		ELSE
			StrSQL = ""
			StrSQL = StrSQL & " INSERT INTO xLog( "
			StrSQL = StrSQL & " LogDate, "
			StrSQL = StrSQL & " PageNum, "
			StrSQL = StrSQL & " WkNum, "
			StrSQL = StrSQL & " UpdtPgCd, "
			StrSQL = StrSQL & " UpdtTime, "
			StrSQL = StrSQL & " UpdtTmnl, "
			StrSQL = StrSQL & " DataCount "
			StrSQL = StrSQL & " )VALUES( "
			StrSQL = StrSQL & "'" & GetYMDStr(Now) & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iNum)) & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iWrkNum)) & "',"
			StrSQL = StrSQL & "'MCOMNINC',"
			StrSQL = StrSQL & "GetDate(),"
			StrSQL = StrSQL & "'" & Session.Contents("userid") & "',"
			StrSQL = StrSQL & "1)"
		End If
		ObjConn.Execute(StrSQL)
	End If

End Function
'2009/03/05 Add-E Tanaka


'日付の表示１
Function DispDateTime(vDate,length)
'	vDate			[ I ]日付
	Dim strTmp
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

%>
