<% 
'本船動静・本船寄港地照会
Function SerchKVsl(conn, rsd, pfp, sSeq, sDate, sName, sPlace, sVsl, sVoy, sBL, sEDate)
    'conn       [ I ]DB接続Object
    'rsd        [ I ]レコードセットObject
    'pfp        [ I ]テンポラリファイル
    'sVsl     	[ I ]船名（コールサイン）
    'sVoy     	[ I ]Voyage番号
    'sBL    	[ I ]BL番号

    bWrite = 0        '出力レコード件数
	bErrFlag = 0

    ' 入力コールサインのチェック
	ConnectSvr conn, rsd
	sql = "SELECT FullName FROM mVessel WHERE VslCode='" & sVsl & "'"

	'SQLを発行してmVesselを検索
	rsd.Open sql, conn, 0, 1, 1
	If rsd.EOF Then
		bErrFlag = 1
	End If
	rsd.Close

	sql = "SELECT VoyCtrl FROM VslSchedule WHERE VslCode='" & sVsl & "' AND DsVoyage='" & sVoy & "'" 

	'SQLを発行してVslScheduleを検索
	rsd.Open sql, conn, 0, 1, 1
	If rsd.EOF Then
		bErrFlag = 1
	Else
		iVoyCtrl = rsd("VoyCtrl")
	End If
	rsd.Close

	If sBL<>"" Then
		sql = "SELECT ShipLine FROM BL WHERE VslCode='" & sVsl & "' And VoyCtrl=" & iVoyCtrl & " And BLNo='" & sBL & "'"

		'SQLを発行してVslScheduleを検索
		rsd.Open sql, conn, 0, 1, 1
		If rsd.EOF Then
			bErrFlag = 1
		End If
		rsd.Close
	End If

	If Not bErrFlag Then
		strOut = sSeq & ","
        strOut = strOut & "IM15,R,"
        strOut = strOut & sDate & ","
        strOut = strOut & sName & ","
        strOut = strOut & sPlace & ","
        strOut = strOut & sVsl & ","
        strOut = strOut & sVoy & ","
        strOut = strOut & sBL & ","
		strOut = strOut & sEDate

        pfp.WriteLine strOut
        bWrite = bWrite + 1

	End If

    SerchKVsl = bWrite    '該当レコードあったかどうか
End Function

'日付のセット
Function SetDateTime(vYear, vMonth, vDay, vHour, vMin)
'	vYear			[ I ]日付(年)
'	vMonth			[ I ]日付(月)
'	vDay			[ I ]日付(日)
'	vHour			[ I ]日付(時)
'	vMin			[ I ]日付(分)
        SetDateTime = trim(vYear) & "/" & _
            Right("0" & trim(vMonth), 2) & "/" & _
            Right("0" & Trim(vDay), 2) & " " & _
            Right("0" & Trim(vHour), 2) & ":" & _
            Right("0" & Trim(vMin), 2)
End Function

'通信日時のセット
Function SetTusinDate()
	'戻り値		[ O ]通信日時(文字列)
				'YYYYMMDDHHNNSS
	SetTusinDate = Trim(Year(Date)) & _
				   Trim(Right("0" & Month(Date), 2)) & _
				   Trim(Right("0" & Day(Date), 2)) & _
				   Trim(Right("0" & Hour(Time), 2)) & _
				   Trim(Right("0" & Minute(Time), 2)) & _
				   Trim(Right("0" & Second(Time), 2))

End Function
'日付文字列の取得
Function GetYMDHM(DateVal)
	'DateVal	[ I ]日付（Date型）
	'戻り値		[ O ]日付文字列（YYYYMMDD形式）

	GetYMDHM = _
		Year(DateVal) & _
		Right("0" & Month(DateVal), 2) & _
		Right("0" & Day(DateVal), 2)   & _
		Right("0" & Hour(DateVal), 2)  & _
		Right("0" & Minute(DateVal), 2)

End Function
'指定年月日のセット２
Function SetTusinDate2(sDate)
'	sDate		[ I ]日付(YYYY/MM/DD HH:MM)
	'戻り値		[ O ]指定日時(YYYYMMDDHHMM)
				' --> 
	SetTusinDate2 = ""
	If sDate <> "" Then
		SetTusinDate2 = Left(Trim(sDate),4) & _
					   Mid(Trim(sDate), 6, 2) & _
					   Mid(Trim(sDate), 9, 2) & _
					   Mid(Trim(sDate), 12, 2) & _
					   Mid(Trim(sDate), 15, 2)
	End If
End Function

'日付の変換処理
Function ChangeDate(sDate)
'	sDate		[ I ]日付(YYYY/MM/DD/HH/MM etc)
	'戻り値		[ O ]指定日時(YYYYMMDDHHMM)
				' --> 
	Dim sDat
	ChangeDate = ""
	If sDate = "////" Then
		Exit Function
	End If
response.write "sDate = "  & sDate
	sChangeWork =  Split(Trim(sDate),"/")

	ChangeDate = sChangeWork(0) & _
				 ArrangeNumV(sChangeWork(1),2) & _
				 ArrangeNumV(sChangeWork(2),2) & _
				 ArrangeNumV(sChangeWork(3),2) & _
				 ArrangeNumV(sChangeWork(4),2)

End Function

' 数値を指定した桁数の文字列に変換(右詰・余白には０)コンテナ入力用
Function ArrangeNumV(nNumber, nFigure)
	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNumV = sNum
End Function

%>