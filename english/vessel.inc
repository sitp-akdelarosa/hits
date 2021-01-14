<% 
'輸出コンテナ照会
Function SerchKCntnr(conn, rsd, pfp, sCont, sSeal, sJyuryo, sSoJyuryo)
    'conn       [ I ]DB接続Object
    'rsd        [ I ]レコードセットObject
    'pfp        [ I ]テンポラリファイル
    'sCont     	[ I ]コンテナNo.
    'sSeal     	[ I ]シールNo.
    'sJyuryo    [ I ]重量

    bWrite = 0        '出力レコード件数

    ' 入力コールサインのチェック
	ConnectSvr conn, rsd
	sql = "SELECT ExportCont.VslCode, ExportCont.VoyCtrl, ExportCont.BookNo, Container.ShipLine, ExportCont.WHArTime, "
	sql = sql & " VslSchedule.LdVoyage "
	sql = sql & " FROM ExportCont, Container, VslSchedule"
	sql = sql & " WHERE ExportCont.ContNo='" & sCont & "' And Container.VslCode = ExportCont.VslCode"
	sql = sql & " AND Container.VoyCtrl = ExportCont.VoyCtrl"
	sql = sql & " AND Container.ContNo = ExportCont.ContNo"
	sql = sql & " AND Container.VslCode = VslSchedule.VslCode"
	sql = sql & " AND Container.VoyCtrl = VslSchedule.VoyCtrl"

	'SQLを発行して輸出コンテナを検索
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
		strOut = Trim(rsd("VslCode")) & ","						' 1:船名
        strOut = strOut & Trim(rsd("LdVoyage")) & ","			' 2:積み次航番号
        strOut = strOut & Trim(sCont) & ","            			' 3:コンテナNo.
        strOut = strOut & Trim(rsd("BookNo")) & ","             ' 4:Booking No.
		strOut = strOut & Trim(rsd("ShipLine")) & ","			' 5:船社
		strOut = strOut & GetYMDHM(rsd("WHArTime")) & ","		' 6:バン詰め日時
        strOut = strOut & sJyuryo & ","         				' 7:重量
        strOut = strOut & sSeal & ","	        				' 8:シールNo.
        strOut = strOut & sSoJyuryo   		      				' 9:総重量

        pfp.WriteLine strOut
        bWrite = bWrite + 1

	End If
	rsd.Close

    SerchKCntnr = bWrite    '該当レコードあったかどうか
End Function

'本船動静・本船寄港地照会
Function SerchKVsl(conn, rsd, pfp, sVsl, sVoy, sPort, sEta, sTa, sTD, sTaL, sTdL)
    'conn       [ I ]DB接続Object
    'rsd        [ I ]レコードセットObject
    'pfp        [ I ]テンポラリファイル
    'sVsl     	[ I ]船名（コールサイン）
    'sVoy     	[ I ]揚げ港／積み港
    'sPort    	[ I ]寄港地
    'sEta     	[ I ]ETA
    'sTa    	[ I ]TA
    'sTD     	[ I ]TD
    'sTaL    	[ I ]ﾛﾝｽｹ入港予定日
    'sTdL    	[ I ]ﾛﾝｽｹ出港予定日

    bWrite = 0        '出力レコード件数

    ' 入力コールサインのチェック
	ConnectSvr conn, rsd
	sql = "SELECT VslSchedule.VslCode, VslSchedule.VoyCtrl, VslSchedule.DsVoyage, VslSchedule.LdVoyage, VslSchedule.ShipLine, "
	sql = sql & " VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.TD, VslPort.ETALong, VslPort.ETDLong"
	sql = sql & " FROM VslSchedule, VslPort"
	sql = sql & " WHERE ( VslSchedule.VslCode='" & sVsl & "' And VslSchedule.DsVoyage = '" & sVoy & "'"
	sql = sql & " AND VslPort.PortCode = '" & sPort & "'"
	sql = sql & " AND VslSchedule.VslCode = VslPort.VslCode"
	sql = sql & " AND VslSchedule.VoyCtrl = VslPort.VoyCtrl ) "
	sql = sql & " Or ( VslSchedule.VslCode='" & sVsl & "' And VslSchedule.LdVoyage = '" & sVoy & "'"
	sql = sql & " AND VslPort.PortCode = '" & sPort & "'"
	sql = sql & " AND VslSchedule.VslCode = VslPort.VslCode"
	sql = sql & " AND VslSchedule.VoyCtrl = VslPort.VoyCtrl ) "

	'SQLを発行して輸出コンテナを検索
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
		strOut = Trim(sVsl) & ","								' 1:船名
        strOut = strOut & Trim(rsd("DsVoyage")) & ","			' 2:揚げ次航番号
        strOut = strOut & Trim(rsd("LdVoyage")) & ","			' 3:積み次航番号
        strOut = strOut & Trim(rsd("ShipLine")) & ","			' 4:船社
        strOut = strOut & Trim(sPort) & ","            			' 5:寄港地名
        strOut = strOut & ChangeDate(sTaL,8) & ","         		' 6:着岸Long Schedule
        strOut = strOut & ChangeDate(sTdL,8) & "," 		        ' 7:離岸Long Schedule
        strOut = strOut & ChangeDate(sEta,12) & ","             	' 8:着岸予定日時
		strOut = strOut & ChangeDate(sTa,12)  & ","				' 9:着岸完了日時
		strOut = strOut & ChangeDate(sTD,12)						' 10:離岸完了日時

        pfp.WriteLine strOut
        bWrite = bWrite + 1

	End If
	rsd.Close

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
Function ChangeDate(sDate,ilength)
'	sDate		[ I ]日付(YYYY/MM/DD/HH/MM etc)
'	ilength		[ I ]長さ　８桁または１２桁
'	戻り値		[ O ]指定日時(YYYYMMDDHHMM)

    Dim sDat
    ChangeDate = ""
    If sDate = "////" Or sDate="" Then
        Exit Function
    End If
    sChangeWork =  Split(Trim(sDate),"/")

    If (ilength = 12 And UBound(sChangeWork)<>4) Or (ilength = 8 And UBound(sChangeWork)<>2) Then
        ChangeDate = "フォーマットが異常です。(" & sDate & ")"
        Exit Function
    End If

    iYear = CInt(sChangeWork(0))
    If iYear<1990 Or iYear>2100 Then
        ChangeDate = "年が異常です。(" & sChangeWork(0) & ")"
        Exit Function
    End If
    iMonth = CInt(sChangeWork(1))
    If iMonth<1 Or iMonth>12 Then
        ChangeDate = "月が異常です。(" & sChangeWork(1) & ")"
        Exit Function
    End If
    iDay = CInt(sChangeWork(2))
    If iDay<1 Or iDay>31 Then
        ChangeDate = "日が異常です。(" & sChangeWork(2) & ")"
        Exit Function
    End If
    If iMonth=4 Or iMonth=6 Or iMonth=9 Or iMonth=11 Then
        If iDay>30 Then
            ChangeDate = "日が異常です。(" & sChangeWork(2) & ")"
            Exit Function
        End If
    End If
    If iMonth=2 Then
        If ( (iYear Mod 4)=0 And (iYear Mod 100)<>0 ) Or (iYear Mod 400)=0 Then
            If iDay>29 Then
                ChangeDate = "日が異常です。(" & sChangeWork(2) & ")"
                Exit Function
            End If
        Else
            If iDay>28 Then
                ChangeDate = "日が異常です。(" & sChangeWork(2) & ")"
                Exit Function
            End If
        End If
    End If
    If ilength>8 Then
        If CInt(sChangeWork(3))<0 Or CInt(sChangeWork(3))>23 Then
            ChangeDate = "時が異常です。(" & sChangeWork(3) & ")"
            Exit Function
        End If
        If CInt(sChangeWork(4))<0 Or CInt(sChangeWork(4))>59 Then
            ChangeDate = "分が異常です。(" & sChangeWork(4) & ")"
            Exit Function
        End If
    End If
	If ilength = 12 Then
		ChangeDate = ArrangeNumV(iYear,4) & _
					 ArrangeNumV(iMonth,2) & _
					 ArrangeNumV(iDay,2) & _
					 ArrangeNumV(sChangeWork(3),2) & _
					 ArrangeNumV(sChangeWork(4),2)
	Else
		ChangeDate = ArrangeNumV(iYear,4) & _
					 ArrangeNumV(iMonth,2) & _
					 ArrangeNumV(iDay,2)
	End If
End Function

' 数値を指定した桁数の文字列に変換(右詰・余白には０)コンテナ入力用
Function ArrangeNumV(nNumber, nFigure)
	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNumV = sNum
End Function

%>