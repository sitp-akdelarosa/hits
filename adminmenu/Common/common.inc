<% 
'SUBDIR="./cam/"
const v_Guide = "GUIDE_BOOK"
const v_Form = "FORM_PATH"
const v_Inform = "INFORM_INI_PATH"
const v_Mail = "MAIL_SEND_FLAG"
const v_Terminal = "TERMINAL_MAP_PATH"
const v_MailTime = "MAIL_SEND_DURATION"
const v_InformUser = "USER_INFORM_UPDATE_DATE"
const v_MailFrom = "MAIL_FROM"
const v_MailSubject = "MAIL_SUBJECT"
const v_MailServer = "MAIL_SERVER"

Dim CRLF
CRLF = Chr(13) & Chr(10)

'-----------------------------
'   HTML Encode
'-----------------------------
function gfHTMLEncode(str) as string
    if gftrim(str)<>"" then
        gfHTMLEncode=server.HTMLEncode(gftrim(str))
    else
        gfHTMLEncode=""
    end if
    return gfHTMLEncode
end function

'-----------------------------
'   SQL Encode ( ' -> '')( \ -> \\)
'----------------------------- 
function gfSQLEncode(str)
    if gftrim(str)<>"" then
        gfSQLEncode=replace(replace(gftrim(str),"'","''"),"\","\\")     '"
    else
        gfSQLEncode=""
    end if
end function

'-----------------------------
'   Trim　NULLの場合→空値(Space0)
'-----------------------------
function gfTrim(str)
    if isnull(str) then
        gfTrim=""
    else
        gfTrim=trim(str)
    end if
end function

'
'ユーザーＩＤが有効かどうかのチェック
'Input :なし
'
Function CheckLoginH()
    ' セッションの有効性をチェック
    Dim strUserID
    strUserID = Trim(session("username"))

    ' セッションが無効なとき
    If strUserID="" Then
      dim strtmp
      strtmp = "ログインしてから"&Session.Timeout&"分経過しました。"
      jampErrerP("0","b000","00","事前情報入力","000",strtmp)
    End If
End Function
'
'エラー画面への遷移
'Input :エラー画面フラグ、画面ID、作業ID、作業名、エラーコード、オプション
'
Function jampErrerP(WinFlag,dispId,wkID,wkName,errerCd,etc)
  on error goto 0		'エラートラップ解除
  Session.Contents("WinFlag")=WinFlag
  Session.Contents("dispId")=dispId
  Session.Contents("wkID")=wkID
  Session.Contents("wkName")=wkName
  Session.Contents("errerCd")=errerCd
  Session.Contents("etc")=etc
  Response.Redirect("./error.asp")
End Function

'ＤＢ接続
Function ConnectSvr(conn, rsd) 
	'conn		[ O ]DB接続Object
	'rsd		[ O ]レコードセットObject

	conn = Server.CreateObject("ADODB.Connection")
	conn.Open "Provider=SQLOLEDB; Data Source = SI-KURT-LAPTOP\SQLEXPRESS; Initial Catalog = CIDB; User Id = sa; Password=P@55w0rd"
	'conn.Open("wavedb", "wave", "wave")

	rsd = Server.CreateObject("ADODB.Recordset")
End Function

'ＤＢ接続
Function ConnectSvrC(conn, rsd) 
	'conn		[ O ]DB接続Object
	'rsd		[ O ]レコードセットObject

	conn = Server.CreateObject("ADODB.Connection")
	conn.Open("camdb", "cam", "cam")

	rsd = Server.CreateObject("ADODB.Recordset")
End Function


'効果確認用データの出力
Function WriteLog(fs, iNum, strScrName, iWrkNum, strOption)
	'fs			[ I ]ファイルシステムObject
	'iNum		[ I ]画面ID
	'strScrName	[ I ]画面名
	'iWrkNum	[ I ]作業ID
	'strOption	[ I ]確認用オプションデータ

	strNowTime = Year(Now) & "/" & _
		Right("0" & Month(Now), 2) & "/" & _
		Right("0" & Day(Now), 2) & " " & _
		Right("0" & Hour(Now), 2) & ":" & _
		Right("0" & Minute(Now), 2) & ":" & _
		Right("0" & Second(Now), 2)
    strDate = GetYMDStr(Now)
    logPath = Server.MapPath("log/" & strDate & ".log")
    logs=fs.OpenTextFile(logPath, 8, True)
    logs.Write(strNowTime)
    logs.Write("," & iNum)
    logs.Write("," & strScrName)
    logs.Write("," & iWrkNum)
    logs.Write("," & Session.Contents("userkind") & "," & Session.Contents("userid"))
    logs.Write("," & Request.ServerVariables("SCRIPT_NAME"))
'   logs.Write("," & Request.ServerVariables("CONTENT_LENGTH"))
    logs.Write("," & Request.ServerVariables("REMOTE_ADDR"))
    If strOption<>"" Then
        logs.WriteLine("," & strOption)
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
        Response.Redirect("userchk.asp?link=" & strReturn)
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

'日付文字列の変換４
Function ChgYMDStr4(sDate)
	'sDate		[ I ]日付文字列（YYYYMMDD形式）
	'戻り値		[ O ]日付文字列（YYYY/MM/DD形式）

	ChgYMDStr4 = Left(sDate, 4) & "/"

	if ( Mid(sDate, 5, 1) = "0" ) Then
		ChgYMDStr4 = ChgYMDStr4 & Mid(sDate, 6, 1) & "/"
	Else
		ChgYMDStr4 = ChgYMDStr4 & Mid(sDate, 5, 2) & "/"
	End If

	if ( Mid(sDate, 7, 1) = "0" ) Then
		ChgYMDStr4 = ChgYMDStr4 & Mid(sDate, 8, 1)
	Else
		ChgYMDStr4 = ChgYMDStr4 & Mid(sDate, 7, 2)
	End If

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
	'vDate			[ I ]日付
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

'日付の表示２
Function DispDateTime2(vDate,length)
'	vDate			[ I ]日付
    If VarType(vDate)<>vbDate Then
        strTmp = ""
    Else
        strTmp = _
            Year(vDate) & "年 " & _
            Right(Month(vDate), 2) & "月" & _
            Right(Day(vDate), 2) & "日 " & _
            Right(Hour(vDate), 2) & ":"

		sDate = Right(Minute(vDate), 2)

        if Len(sDate) = 1 Then
            strTmp = strTmp & "0" & sDate
        Else
            strTmp = strTmp & sDate
        End If

        DispDateTime2 = strTmp
    End If
End Function

'ヘッダーの描画
Sub DisplayHeader()
	Response.Write("<tr bgcolor='#0066CC'>")
	Response.Write("<td height='7' colspan='2' valign='top'><img src='/images/transparent.gif' width='1' height='1'></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td height='1' colspan='2' valign='top' bgcolor='#FFFFFF'><img src='/images/transparent.gif' width='1' height='1'></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td height='75' colspan='2' valign='top'><img src='/images/topbanner1.jpg' width='830' height='75'></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td height='2' colspan='2' valign='top'><img src='/images/transparent.gif' width='1' height='1'></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td height='5' colspan='2' valign='top' bgcolor='#0066CC'><img src='/images/transparent.gif' width='1' height='1'></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td align='right' width='100%' height='48'></td>")
	Response.Write("</tr>")
End Sub

'ヘッダーの描画
Sub DisplayHeader2(strTitle)
	Response.Write("<tr bgcolor='#0066CC'><td height='5' valign='top'><img src='/images/transparent.gif' width='1' height='1'></td></tr>")
    Response.Write("<tr><td height='1' valign='top' bgcolor='#FFFFFF'><img src='/images/transparent.gif' width='1' height='1'></td></tr>")
    Response.Write("<tr><td height='75'  valign='middle' align='center' class='headermenu'>" & strTitle & "</td></tr>")
    Response.Write("<tr><td height='1'  valign='top'><img src='/images/transparent.gif' width='1' height='1'></td></tr>")
    Response.Write("<tr><td height='5'  valign='top' bgcolor='#0066CC'><img src='/images/transparent.gif' width='1' height='1'></td></tr>")
    Response.Write("<tr><td align='right' width='100%' height='48'></td></tr>")
End Sub


'フッターの描画
Sub DisplayFooter
	Response.Write("<tr bgcolor='#3366CC'>")
	Response.Write("<td height='15' colspan='3' class='footer2'>Internet Explorer 5.5 SP2・NetScape6.0　以降のブラウザでご覧ください</td>")
	Response.Write("</tr>")
End Sub

'Menuバーの描画
Sub DispMenuBar
    Response.Write("<table width='100%' border='0' cellspacing='0' cellpadding='0'>")
    Response.Write("  <tr>")
    Response.Write("    <td valign='bottom' align='right'><a href='http://www.hits-h.com/index.asp'><img src='/gif/b-home.gif' border='0' width='270' height='23' usemap='#map'></a></td>")
    Response.Write("  </tr>")
    Response.Write("  <tr>")
    Response.Write("    <td bgcolor='000099' height='10'><img src='/gif/1.gif'></td>")
    Response.Write("  </tr>")
    Response.Write("</table>")
End Sub

'エラーメッセージの表示
Sub DispErrorMessage(strMessage)
    Response.Write("<center>")
    Response.Write("<table border=1 cellpadding='2' cellspacing='1'>")
    Response.Write("<tr>")
    Response.Write("<td bgcolor='#FFFFFF'>")
    Response.Write("<table border='0'>")
    Response.Write("<tr>")
    Response.Write("<td valign='middle'><img src='gif/error.gif'></td>")
    Response.Write("<td><b><font color='#FF4400'>")
    Response.Write(strMessage)
    Response.Write("</font></b></td>")
    Response.Write("</tr>")
    Response.Write("</table>")
    Response.Write("</td>")
    Response.Write("</tr>")
    Response.Write("</table>")
    Response.Write("</center>")
End Sub

'注目メッセージの表示
Sub DispInformationMessage(strMessage)
    Response.Write("<center>")
    Response.Write("<table border=1 cellpadding='2' cellspacing='1'>")
    Response.Write("<tr>")
    Response.Write("<td bgcolor='#FFFFFF'>")
    Response.Write("<table border='0'>")
    Response.Write("<tr>")
    Response.Write("<td valign='middle'><img src='gif/info.gif'></td>")
    Response.Write("<td><b><font color='#0000FF'>")
    Response.Write(strMessage)
    Response.Write("</font></b></td>")
    Response.Write("</tr>")
    Response.Write("</table>")
    Response.Write("</td>")
    Response.Write("</tr>")
    Response.Write("</table>")
    Response.Write("</center>")
End Sub

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




'配列が正確な日付フォーマットかどうか
Function IsRightDateFormat(strary)
	result = False

	If UBound(strary)=1 Then
		If IsNumeric(strary(0)) and IsNumeric(strary(1)) Then
			If (strary(0)>2000) And (strary(0)<2100) And (strary(1)>0) And (strary(1)<13) Then
				result = True
			End If
		End If
	End If
	If UBound(strary)=2 Then
		If IsNumeric(strary(0)) and IsNumeric(strary(1)) and IsNumeric(strary(2)) Then
			If (strary(0)>2000) And (strary(0)<2100) And (strary(1)>0) And (strary(1)<13) And (strary(2)>0) And (strary(2)<32) Then
				result = True
			End If
		End If
	End If

	IsRightDateFormat = result
End Function

'次月の1日を日付フォーマットで返す
Function FirstDayOfNextMonth(yearstr, monthstr)
	If monthstr="12" Then
		monthstr = "1"
		yearstr = yearstr + 1
	Else
		monthstr = monthstr + 1
	End If

	FirstDayOfNextMonth = yearstr & "/" & monthstr & "/1"
End Function

'SQLフォーマットに変換
Function ToSQLFormat(str)
	If IsNull(str) Then
		ToSQLFormat = str
	Else
		result = ""

		For i=1 To Len(str)
			c = Mid(str,i,1)

			If c="'" Then
				result = result & "'"
			End If

			result = result & c
		Next

		ToSQLFormat = result
	End If
End Function

'指定文字が1バイトかどうか
Function Is1ByteChar(c)
	If (Asc(c)>=&H20 and Asc(c)<=&H7e) or (Asc(c)>=161 and Asc(c)<=223) Then
		Is1ByteChar = True
	Else
		Is1ByteChar = False
	End If
End Function

'文字列を指定桁数で切る
Function FormatOptionalDigit(str, digit)
	If IsNull(str) Then
		FormatOptionalDigit = str
	Else
		result = ""
		count = 0
		For i=1 To Len(str)
			c = Mid(str,i,1)

			If Is1ByteChar(c)=True Then
				count = count + 1
			Else
				count = count + 2
			End If

			If count<=digit Then
				result = result & c
			Else
				Exit For
			End If
		Next

		FormatOptionalDigit = result
	End If
End Function

'指定桁数毎に改行タグを入れる
Function InsertReturnCodeAtEveryOptionalDigit(str, digit)
	If IsNull(str) Then
		InsertReturnCodeAtEveryOptionalDigit = str
	Else
		result = ""
		count = 0
		For i=1 To Len(str)
			c = Mid(str,i,1)

			If Is1ByteChar(c)=True Then
				count = count + 1
			Else
				count = count + 2
			End If

			If count>digit Then
				count = 0
				result = result & "<br>"
			End If

			result = result & c
		Next

		InsertReturnCodeAtEveryOptionalDigit = result
	End If
End Function

'数値3桁毎にカンマを入れる
Function ToNumberFormat(str)
	If IsNull(str) Then
		ToNumberFormat = str
	Else
		result = str
		commaindex = InStrRev(str, ".")

		If commaindex>0 Then
			result = Left(str, commaindex-1)
		End If

		tmp = ""
		revtmp = StrReverse(result)
		For i=1 To Len(result)
			tmp = tmp & Mid(revtmp,i,1)
			If (i Mod 3=0) and (i<Len(result)) Then
				tmp = tmp & ","
			End If
		Next
		result = StrReverse(tmp)

		If commaindex>0 Then
			size = Len(str) - commaindex + 1
			result = result & Right(str, size)
			If size=2 Then
				result = result & "0"
			End If
		Else
			result = result & ".00"
		End If

		ToNumberFormat = result
	End If
End Function

%>