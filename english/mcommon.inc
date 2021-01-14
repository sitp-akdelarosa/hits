<%
' ユーザIDの形式チェック
Function CheckUserID(sUserID)
	'sUserID		[ I ]ユーザID
	'CheckUserID	[ O ]エラーメッセージ

	' 運送会社会社コード(ユーザIDの先頭)の最大値
	Const COMPANY_MAX = "200"

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
			If IsNumeric(strKeyName) Then
				GetKeyLabel = "&#" & CStr(63878 + CInt(strKeyName))
			Else
				GetKeyLabel = "[" & Right(strKeyName,1) & "]"
			End If
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
	'CanCarryOut	[ O ]搬出可能か

	Dim tNow
	tNow = Now
	
	' 通関／保税輸送のチェック
	If Trim(rs("OLTICFlag")) = "I" Then
		If Trim(rs("OLTICNo")) = "" Then	
			CanCarryOut = False
			Exit Function
		End If
	Else
		If Not (tNow >= rs("OLTDateFrom") And tNow <= rs("OLTDateTo")) Then
			CanCarryOut = False
			Exit Function
		End If		
	End If

	If Trim(rs("DOStatus")) = "Y" And _
		(rs("FreeTime") >= tNow Or rs("FreeTimeExt") >= tNow) And _
		rs("DelPermitDate") <= tNow Then
		
		CanCarryOut = True
	Else
		CanCarryOut = False
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
Function WriteLogM(fs, strUserID, strScrName, strOption)
	'fs			[ I ]ファイルシステムObject
	'strUserID	[ I ]ユーザID
	'strScrName	[ I ]画面名
	'strOption	[ I ]確認用オプションデータ

	' Logを出力するフォルダ
	Const LOG_FOLDER = "../log"

	' Logファイルの拡張子 
	Const LOG_EXTENT = "log"

	Dim strDate
	Dim logPath
	Dim logs

	strDate = GetYMDStr(Now)
	logPath = Server.MapPath(LOG_FOLDER & "/" & strDate & "." & LOG_EXTENT)
	Set logs=fs.OpenTextFile(logPath, 8, True)
	logs.Write Now
	logs.Write Chr(9) & strScrName
	logs.Write Chr(9) & strUserID
	logs.Write Chr(9) & Request.ServerVariables("SCRIPT_NAME")
'	logs.Write Chr(9) & Request.ServerVariables("CONTENT_LENGTH")
	logs.Write Chr(9) & Request.ServerVariables("REMOTE_ADDR")
	If strOption<>"" Then
		logs.WriteLine Chr(9) & strOption
	Else
		logs.WriteLine 
	End If

	logs.Close
End Function
%>
