<%
'輸出照会関係関数

'輸出コンテナ照会
Function SerchExpCntnr(conn, rsd, pfp, sWhere)
    'conn       [ I ]DB接続Object
    'rsd        [ I ]レコードセットObject
    'pfp        [ I ]テンポラリファイル
    'sWhere     [ I ]検索条件

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    bWrite = 0        '出力レコード件数

    sql = "SELECT ExportCont.BookNo, ExportCont.ContNo, ExportCont.VslCode, ExportCont.PickPlace, " & _
          "ExportCont.RFFlag, ExportCont.VoyCtrl, ExportCont.EmpDelTime, ExportCont.WHArTime, " & _
          "ExportCont.VanTime, ExportCont.ExpArTime, ExportCont.OutTime, ExportCont.SYRecTime, " & _
          "ExportCont.CYRecPrDate, ExportCont.SYRecFlag, ExportCont.SYPRRecTime, ExportCont.DGFlag, " & _
          "Container.ContSize, Container.ContType, Container.ContHeight, Container.SealNo, Container.ContWeight, " & _
          "Container.CargoWeight, Container.ShipLine, " & _
          "Booking.RecTerminal, Booking.DPort, Booking.RFFlag BRFFlag, Booking.DGFlag BDGFlag, " & _
          "ExportCargoInfo.WHArTime WHArSchDate, ExportCargoInfo.CYRecDate, " & _
          "VslSchedule.LdVoyage, VslSchedule.CYOpen, VslSchedule.CYCut, " & _
          "mShipLine.FullName ShipLineName, " & _
          "mVessel.FullName ShipName, " & _
          "VslPort.ETD, VslPort.TD, VslPort.ETDLong, Container.TareWeight "
'☆☆☆ Add_S  by nics 2010.02.22
    sql = sql &  ", ExportCont.CustOK "
'☆☆☆ Add_E  by nics 2010.02.22
'☆☆☆ Add_S  by nics 2009.03.05
    sql = sql &  ", InspectCont.XRStatus, InspectCont.XDStatus "
' - NICS不具合(2009/07/22)    sql = sql &  ", (Select TOP 1 StockPlace From mPlaceCode Where mPlaceCode.Operator=Booking.RecTerminalCode) As StockPlace "
    sql = sql &  ", (Select TOP 1 StockPlace From mPlaceCode Where mPlaceCode.Operator=(Select TOP 1 mSL.OpeCode From mShipLine mSL Where mSL.ShipLine=VslSchedule.ShipLine) and mPlaceCode.Terminal=Booking.RecTerminalCode) As StockPlace "
    sql = sql &  ", (Select TOP 1 mOperator.FullName From mShipLine mSL, mOperator Where mSL.ShipLine=VslSchedule.ShipLine AND mOperator.OpeCode=mSL.OpeCode) As OpeName "
'☆☆☆ Add_E  by nics 2009.03.05
    sql = sql & " FROM ExportCont, Container, Booking, ExportCargoInfo, VslSchedule, mShipLine, mVessel, VslPort "
'☆☆☆ Add_S  by nics 2009.03.05
    sql = sql &  ", InspectCont "
'☆☆☆ Add_E  by nics 2009.03.05
    sql = sql & "WHERE (" & sWhere & ") AND Container.VslCode=ExportCont.VslCode AND Container.VoyCtrl=ExportCont.VoyCtrl AND " & _
          "Container.ContNo=ExportCont.ContNo AND Booking.VslCode=*ExportCont.VslCode AND Booking.VoyCtrl=*ExportCont.VoyCtrl AND " & _
          "Booking.BookNo=*ExportCont.BookNo AND ExportCargoInfo.ContNo=*ExportCont.ContNo AND ExportCargoInfo.BookNo=*ExportCont.BookNo AND ExportCargoInfo.VslCode=*ExportCont.VslCode AND " & _
          "VslSchedule.VslCode=*ExportCont.VslCode AND VslSchedule.VoyCtrl=*ExportCont.VoyCtrl AND " & _
          "mShipLine.ShipLine=*Container.ShipLine AND mVessel.VslCode=*ExportCont.VslCode AND " & _
          "VslPort.VslCode=*ExportCont.VslCode AND VslPort.VoyCtrl=*ExportCont.VoyCtrl AND VslPort.PortCode='" & strPort & "'"
'☆☆☆ Add_S  by nics 2009.03.05
    sql = sql &  " AND InspectCont.VslCode=*ExportCont.VslCode AND InspectCont.VoyCtrl=*ExportCont.VoyCtrl AND " & _
          "InspectCont.ContNo=*ExportCont.ContNo " 
'☆☆☆ Add_E  by nics 2009.03.05
    sql = sql & " ORDER BY ExportCont.BookNo, ExportCont.ContNo"

    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1

    Dim sOutText()

    Do While Not rsd.EOF
        ' 出力レコードの編集
        strOut = Trim(rsd("BookNo")) & ","                     ' 1:Booking No.
        strOut = strOut & Trim(rsd("ContNo")) & ","            ' 2:コンテナNo.
        strOut = strOut & Trim(rsd("PickPlace")) & ","         ' 3:空コン受取場所
        strOut = strOut & Trim(rsd("ContSize")) & ","          ' 4:サイズ
        strOut = strOut & Trim(rsd("ContHeight")) & ","        ' 5:高さ
        ' 6:リーファー
        strRFFlag = Trim(rsd("RFFlag"))
        If IsNull(strRFFlag) Or strRFFlag="" Then
            strRFFlag = Trim(rsd("BRFFlag"))
        End If
        If IsNull(strRFFlag) Or strRFFlag="" Then
            If DispDateTime(rsd("VanTime"),0)<>"" Or DispDateTime(rsd("ExpArTime"),0)<>"" Then
				strRFFlag = "-"
            End If
        End If
		strOut = strOut & strRFFlag & ","
        strOut = strOut & Trim(rsd("RecTerminal")) & ","       ' 7:搬入ターミナル名
        strOut = strOut & Trim(rsd("SealNo")) & ","            ' 8:シールNo.
        strOut = strOut & Trim(rsd("ContWeight")) & ","        ' 9:総重量
        strOut = strOut & DispDateTime(rsd("CYOpen"),10) & "," '10:ＣＹオープン
        strOut = strOut & DispDateTime(rsd("CYCut"),10) & ","  '11:ＣＹクローズ

        If IsNull(rsd("ShipLineName")) Then
            strOut = strOut & Trim(rsd("ShipLine")) & ","      '12:船社名
        Else
            strOut = strOut & Trim(rsd("ShipLineName")) & ","  '12:船社名
        End If
        If IsNull(rsd("ShipName")) Then
            strOut = strOut & Trim(rsd("VslCode")) & ","       '13:船名
        Else
            strOut = strOut & Trim(rsd("ShipName")) & ","      '13:船名
        End If
        strOut = strOut & Trim(rsd("LdVoyage")) & ","          '14:次航
        strOut = strOut & Trim(rsd("DPort")) & ","             '15:仕向港(コード)

        ' 時系列のチェック
        strEmpDelTime = DispDateTime(rsd("EmpDelTime"),0)             '17:空コン受取時刻
        strWHArTime = DispDateTime(rsd("WHArTime"),0)                 '18:倉庫到着
        strVanTime = DispDateTime(rsd("VanTime"),0)                   '19:バンニング
        strSYRecTime = DispDateTime(rsd("SYRecTime"),0)               '25:ストックヤード搬入時刻
        strExpArTime = DispDateTime(rsd("ExpArTime"),0)               '20:ＣＹ搬入
        strOutTime = DispDateTime(rsd("OutTime"),0)                   '21:船積完了
        If strWHArTime<strEmpDelTime Then
            strWHArTime = ""                    ' 倉庫到着が空コン受取時刻より過去のとき表示しない
        End If
        strFuture = strSYRecTime
        If strFuture="" Then
            strFuture = strExpArTime
        End If
        If strFuture="" Then
            strFuture = strOutTime
        End If
        If strFuture<>"" And strWHArTime>strFuture Then
            strWHArTime = ""                    ' 倉庫到着がKACCS,Hits時刻より未来のとき表示しない
        End If
        strOld = strWHArTime
        If strOld="" Then
            strOld = strEmpDelTime
        End If
        If strVanTime<strOld Then
            strVanTime = ""                     ' バンニングが倉庫到着または空コン受取時刻より過去のとき表示しない
        End If
        If strFuture<>"" And strVanTime>strFuture Then
            strVanTime = ""                     ' バンニングがKACCS,Hits時刻より未来のとき表示しない
        End If

        strCheck=rsd("ShipName")                '船名がなかったら、離岸・着岸時刻を表示しないようにする

	'船名のチェックを止める(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ","
'        Else
            strOut = strOut & DispDateTime(rsd("ETD"),0) & ","        '16:離岸予定時刻(ターミナル)
'        End If
	'船名のチェックを止める(2004/03/12)
        strOut = strOut & DispDateTime(rsd("EmpDelTime"),0) & ","     '17:空コン受取時刻
        strOut = strOut & strWHArTime & ","                           '18:倉庫到着
        strOut = strOut & strVanTime & ","                            '19:バンニング
        strOut = strOut & DispDateTime(rsd("ExpArTime"),0) & ","      '20:ＣＹ搬入
        strOut = strOut & DispDateTime(rsd("OutTime"),0) & ","        '21:船積完了
	'船名のチェックを止める(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ",,,"
'        Else
            strOut = strOut & DispDateTime(rsd("TD"),0) & ","         '22:離岸完了時刻(ターミナル)
            strOut = strOut & ",,"
'            strOut = strOut & DispDateTime(rsd("TA"),0) & ","         '23:着岸完了時刻(仕向港)
'            strOut = strOut & DispDateTime(rsd("ETA"),0) & ","        '24:着岸予定時刻(仕向港)
'        End If
	'船名のチェックを止める(2004/03/12)

        strOut = strOut & DispDateTime(rsd("SYRecTime"),0) & ","      '25:ストックヤード搬入時刻
	'船名のチェックを止める(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ","
'        Else
            strOut = strOut & DispDateTime(rsd("ETDLong"),0) & ","    '26:離岸スケジュール時刻(ターミナル)
'        End If
	'船名のチェックを止める(2004/03/12)
        strOut = strOut & DispDateTime(rsd("WHArSchDate"),0) & ","    '27:空倉庫到着スケジュール時刻

        strOut = strOut & Trim(rsd("CargoWeight")) & ","              '28:貨物重量
        strOut = strOut & Trim(rsd("SYRecFlag")) & ","                '29:ストックヤード搬入予約フラグ
        strOut = strOut & DispDateTime(rsd("SYPRRecTime"),0) & ","    '30:ストックヤード搬入予約日時
        If IsNull(rsd("CYRecPrDate")) Then
            strOut = strOut & DispDateTime(rsd("CYRecDate"),0) & ","  '31:ＣＹ搬入指定日時
        Else
            strOut = strOut & DispDateTime(rsd("CYRecPrDate"),0) & ","'31:ＣＹ搬入指定日時
        End If
        '32:危険品
        strDGFlag = Trim(rsd("DGFlag"))
        If IsNull(strDGFlag) Or strDGFlag="" Then
            strDGFlag = Trim(rsd("BDGFlag"))
        End If
        If IsNull(strDGFlag) Or strDGFlag="" Then
            If DispDateTime(rsd("VanTime"),0)<>"" Or DispDateTime(rsd("ExpArTime"),0)<>"" Then
				strDGFlag = "-"
            End If
        End If
'--- mod by mes(2005/3/28)
'		strOut = strOut & strDGFlag
		strOut = strOut & strDGFlag & ","
'--- end mes

'--- add by mes(2005/3/28)
		' 33:テアウェイト
        strOut = strOut & Trim(rsd("TareWeight"))
'		If Trim(rsd("TareWeight"))<>"" Then
'	        strOut = strOut & (Trim(rsd("TareWeight")) * 1000)
'	    Else
'	    	strOut = strOut & Trim(rsd("TareWeight"))
'	    End If
'--- end mes

'☆☆☆ Add_S  by nics 2009.03.05
'☆☆☆ Add_S  by nics 2010.03.23	英語モードに上書き編集
        ' 34:X線有無
        If Trim(rsd("XRStatus")) <> "" then
            strOut = strOut & "," & "Need"		
        Else					
            strOut = strOut & "," & "Not need"	
        End If

        ' 35:X線CY返却
        Select Case Trim(rsd("XRStatus"))
            Case "S"                            ' 対象		有
                strOut = strOut & "," & "Will be returned"
            Case "O"                            ' 検査場	搬出中
                strOut = strOut & "," & "CY Out(will be returned)"
            Case "C"                            ' 完了		済
                strOut = strOut & "," & "Returned"
            Case Else                           ' 他		−
                strOut = strOut & "," & "Other"
        End Select

        ' 36:ターミナル搬入確認
        If Trim(rsd("ExpArTime")) <> "" then		'済
            strOut = strOut & "," & "Not arrived"
        Else						'未
            strOut = strOut & "," & "Arrived"
        End If
'☆☆☆ Add_E  by nics 2010.03.23	英語モードに上書き編集

        ' 37:蔵置場所コード
        strOut = strOut & "," & Trim(rsd("StockPlace"))

        ' 38:担当オペレータ
        strOut = strOut & "," & Trim(rsd("OpeName"))
'☆☆☆ Add_E  by nics 2009.03.05
'☆☆☆ Add_S  by nics 2010.02.22＆英語モードに上書き編集
        ' 38:輸出通関情報(数字のコメントが間違っているので38のみ修正した
        If Trim(rsd("CustOK")) = "Y" then
            strOut = strOut & "," & "Cleared"
        Else
            strOut = strOut & "," & "Not cleared"
        End If
'☆☆☆ Add_E  by nics 2010.02.22＆英語モードに上書き編集

' 39:コンテナタイプ
'Add-S MES Aoyagi 2010.11.23 コンテナタイプを表示する。
	strOut = strOut & "," & Trim(rsd("ContType"))
'Add-E MES Aoyagi 2010.11.23 コンテナタイプを表示する。

        ReDim Preserve sOutText(bWrite)
        sOutText(bWrite) = strOut
        bWrite = bWrite + 1

        rsd.MoveNext
    Loop

    rsd.Close

    For i=0 To bWrite-1
        strTmp=Split(sOutText(i),",")

        ' 仕向港データを探す
        If strTmp(14)<>"" Then
            sql = "SELECT ExportCont.VoyCtrl, VslPort.TA, VslPort.ETA FROM ExportCont, VslPort " & _
                  "WHERE ExportCont.ContNo='" & strTmp(1) & "' AND ExportCont.BookNo='" & strTmp(0) & "' AND VslPort.VslCode=ExportCont.VslCode AND " & _
                  "VslPort.VoyCtrl=ExportCont.VoyCtrl AND VslPort.PortCode='" & strTmp(14) & "'"
            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                If strTmp(12)<>"" Then    '船名があったら
                    strTmp(22) = DispDateTime(rsd("TA"),0)     '23:着岸完了時刻(仕向港)
                    strTmp(23) = DispDateTime(rsd("ETA"),0)    '24:着岸予定時刻(仕向港)
                End If
            End If
            rsd.Close

            sql = "SELECT mPort.FullName From mPort " & _
                  "WHERE mPort.PortCode='" & strTmp(14) & "'"
            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                If Not IsNull(rsd("FullName")) Then
                    strTmp(14) = Trim(rsd("FullName"))         '15:仕向港
                End If
            End If
            rsd.Close

            strWork=""
            For j=0 To UBound(strTmp)
                strWork=strWork & strTmp(j) & ","
            Next
            sOutText(i)=strWork
        End If

        pfp.WriteLine sOutText(i)
    Next

    SerchExpCntnr = bWrite    '該当レコードあったかどうか
End Function

'輸出コンテナCSVファイルタイトル行出力
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]検索モード(true=コンテナ検索 / false=ブッキング検索)

    Response.Write "Booking_No.,"
    Response.Write "Container_No.,"
    Response.Write "Pickup_place,"
    Response.Write "Size,"
'Add-S MES Aoyagi 2010.11.23 コンテナタイプも表示 
    Response.Write "Type,"
'Add-E MES Aoyagi 2010.11.23 コンテナタイプも表示
    Response.Write "Height,"
    Response.Write "Reefer,"
    Response.Write "Seal_No.,"
    Response.Write "CW,"
    Response.Write "GW,"
    Response.Write "Shipping_Yard,"
    Response.Write "CY_open,"
    Response.Write "CY_close,"

    Response.Write "EmptyContainer_PickupTime,"
    Response.Write "Warehouse_ArrivalTime(Instruction),"
    Response.Write "Warehouse_ArrivalTime(Actual),"
    Response.Write "Vanning_Time,"
    Response.Write "SY_ReservationTime,"
    Response.Write "SY_InTime(Actual),"
    Response.Write "CY_InTime(Instruction),"
    Response.Write "CY_InTime(Actual),"
    Response.Write "Loading_Time(Actual),"
    Response.Write "Departure_Time(Estimate),"
    Response.Write "Departure_Time(Intended),"
    Response.Write "Departure_Time(Actual),"
    Response.Write "DischargePort_ArrivalTime(Intended),"
    Response.Write "DischargePort_ArrivalTime(Actual),"

    Response.Write "Ship's_Line,"
    Response.Write "Vessel_Name,"
    Response.Write "Voyage No.,"
    Response.Write "Discharge_Port,"
    Response.Write "Hazard,"
'--- add by mes(2005/3/28)
	' テアウェイト追加
    Response.Write "TW"
'--- end mes

'☆☆☆ Add_S  by nics 200902改造
    Response.Write ",X-ray_Inspection(Need)"
    Response.Write ",X-ray_Inspection(Return)"
    Response.Write ",ShippingYard(code)"
    Response.Write ",Operater"
'☆☆☆ Add_S  by nics 201002改造
    Response.Write ",Clearance"
'☆☆☆ Add_E  by nics 201002改造
    Response.Write ",ConfirmationOfArrival"
'☆☆☆ Add_E  by nics 200902改造

    Response.Write Chr(13) & Chr(10)
End Sub

'輸出コンテナCSVファイルデータ行出力
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]検索モード(true=コンテナ検索 / false=ブッキング検索)
    'pfp        [ I ]テンポラリファイル

    ' ダウンロードファイルのレコードがある間繰り返す
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        ' Booking No.
        If Not bMode Then
            Response.Write anyTmp(0) & ","
        Else
            Response.Write ","
        End If
        ' コンテナNo.
        Response.Write anyTmp(1) & ","

        ' 空コン受取場所
        Response.Write anyTmp(2) & ","
        ' サイズ
        Response.Write anyTmp(3) & ","
'Add-S MES Aoyagi 2010.11.23 コンテナタイプも表示
	'タイプ	
	Response.Write anyTmp(39) & ","
'Add-E MES Aoyagi 2010.11.23 コンテナタイプも表示
        ' 高さ
        Response.Write anyTmp(4) & ","
       ' リーファー
        If anyTmp(5)="R" Then
            Response.Write "○,"
        ElseIf anyTmp(5)<>"" Then
            Response.Write "−,"
        Else
            Response.Write ","
        End If
        ' シールNo.
        Response.Write anyTmp(7) & ","
        ' 貨物重量
        If anyTmp(27)<>"" And anyTmp(27)<>"0" Then
            dWeight=anyTmp(27) / 10
            Response.Write dWeight
        Else
            Response.Write "−"
        End If
        Response.Write ","
        ' 総重量
        If anyTmp(8)<>"" And anyTmp(8)<>"0" Then
            dWeight=anyTmp(8) / 10
            Response.Write dWeight
        Else
            Response.Write "−"
        End If
        Response.Write ","
        ' 搬入ターミナル名
        Response.Write anyTmp(6) & ","
        ' ＣＹオープン
        Response.Write anyTmp(9) & ","
        ' ＣＹクローズ
        Response.Write anyTmp(10) & ","

        ' 空コン受取時刻
        Response.Write anyTmp(16) & ","
        ' 倉庫到着スケジュール
        Response.Write anyTmp(26) & ","
        ' 倉庫到着
        Response.Write anyTmp(17) & ","
        ' バンニング
        Response.Write anyTmp(18) & ","
        ' ストックヤード搬入予約
        sTemp=DispReserveCell(anyTmp(28),anyTmp(29),sColor)
        If anyTmp(28)>="1" And anyTmp(28)<="4" Then
            sTemp=Left(anyTmp(29),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' ストックヤード搬入
        Response.Write anyTmp(24) & ","
        ' ＣＹ搬入指示
        Response.Write anyTmp(30) & ","
        ' ＣＹ搬入
        Response.Write anyTmp(19) & ","
        ' 船積完了
        Response.Write anyTmp(20) & ","
        ' 離岸スケジュール
        Response.Write anyTmp(25) & ","
        ' 離岸予定時刻
        Response.Write anyTmp(15) & ","
        ' 離岸完了時刻
        Response.Write anyTmp(21) & ","
        ' 着岸予定時刻
        Response.Write anyTmp(23) & ","
        ' 着岸完了時刻
        Response.Write anyTmp(22) & ","

        ' 船社名
        Response.Write anyTmp(11) & ","
        ' 船名
        Response.Write anyTmp(12) & ","
        ' 次航
        Response.Write anyTmp(13) & ","
        ' 仕向港
        Response.Write anyTmp(14) & ","

        ' 危険品
        If anyTmp(31)="H" Then
            Response.Write "○"
        ElseIf anyTmp(31)<>"" Then
            Response.Write "−"
        Else
            Response.Write ""
        End If

'--- add by mes(2005/3/28)
        Response.Write ","

        ' テアウェイト
        If anyTmp(32)<>"" And anyTmp(32)>0 Then
	    	If anyTmp(32)<100 then
		        dWeight=anyTmp(32) * 100
		    Else
		        dWeight=anyTmp(32)
		    End If
            Response.Write dWeight
        Else
            Response.Write "−"
        End If
'--- end mes

'☆☆☆ Add_S  by nics 200902改造
        'X線検査有無
        Response.Write "," & anyTmp(33)
        'X線検査CY返却
        Response.Write "," & anyTmp(34)
        '蔵置場所コード
        Response.Write "," & anyTmp(36)
        '担当オペレータ
        Response.Write "," & anyTmp(37)
'☆☆☆ Add_S  by nics 201002改造
        '税関
        Response.Write "," & anyTmp(38)
'☆☆☆ Add_E  by nics 201002改造
        'ターミナル搬入確認
        Response.Write "," & anyTmp(35)
'☆☆☆ Add_E  by nics 200902改造

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>