<%
'輸出照会関係関数

'輸出コンテナ照会
Function SerchMSExpCntnr(conn, rsd, pfp, sWhere, sOrder)
    'conn       [ I ]DB接続Object
    'rsd        [ I ]レコードセットObject
    'pfp        [ I ]テンポラリファイル
    'sWhere     [ I ]検索条件
    'sOrdr      [ I ]Sort条件

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    bWrite = 0        '出力レコード件数

    sql = "SELECT ExportCargoInfo.Shipper, ExportCargoInfo.ShipCtrl, ExportCargoInfo.VslCode, " & _
          "ExportCargoInfo.LdVoyage, ExportCargoInfo.BookNo, " & _
          "ExportCargoInfo.Forwarder, ExportCargoInfo.ContNo, ExportCargoInfo.Trucker, " & _
          "ExportCargoInfo.WHArTime, ExportCargoInfo.CYRecDate, " & _
          "ExportCargoInfo.ContSize, ExportCargoInfo.ContType, " & _
          "ExportCargoInfo.ContHeight, ExportCargoInfo.PickPlace, " & _
          "ExportCargoInfo.Remark, ExportCargoInfo.TruckRefFlag, " & _
          "ExportCargoInfo.PickDate, ExportCargoInfo.OpeCode, " & _
          "ExportCargoInfo.PickOKFlag, ExportCargoInfo.PkPlUpdFlag, " & _
          "ExportCargoInfo.PkDtUpdFlag, " & _
          "VslSchedule.ShipLine, " & _
          "VslSchedule.CYCut, " & _
          "ExportCont.VoyCtrl, " & _
          "Booking.VoyCtrl BookVoyage, " & _
          "Booking.RecTerminal, " & _
          "mVessel.FullName ShipFullName, mVessel.NameAbrev ShipName, " & _
          "mOperator.NameAbrev OpeName, " & _
          "mShipper.NameAbrev ShipperName, mShipper.FullName ShipperFullName, " & _
          "mForwarder.NameAbrev ForName, mForwarder.FullName ForFullName, " & _
          "mTrucker.NameAbrev TruckName, mTrucker.FullName TruckFullName"
    sql = sql & " FROM ExportCargoInfo, VslSchedule, ExportCont, Booking, mVessel, mOperator, mShipper, mForwarder, mTrucker "
    sql = sql & "WHERE (" & sWhere & ") AND VslSchedule.VslCode=*ExportCargoInfo.VslCode AND VslSchedule.LdVoyage=*ExportCargoInfo.LdVoyage AND " & _
          "ExportCont.VslCode=*ExportCargoInfo.VslCode AND ExportCont.BookNo=*ExportCargoInfo.BookNo AND ExportCont.ContNo=*ExportCargoInfo.ContNo AND " & _
          "Booking.VslCode=*ExportCargoInfo.VslCode AND Booking.BookNo=*ExportCargoInfo.BookNo AND " & _
          "mVessel.VslCode=*ExportCargoInfo.VslCode AND " & _
          "mShipper.Shipper=*ExportCargoInfo.Shipper AND " & _
          "mOperator.OpeCode=*ExportCargoInfo.OpeCode AND " & _
          "mForwarder.Forwarder=*ExportCargoInfo.Forwarder AND " & _
          "mTrucker.Trucked=*ExportCargoInfo.Trucker"
    sql = sql & " ORDER BY " & sOrder

    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1

    Do While Not rsd.EOF
        sText=""

        ' 出力レコードの編集
        sText = Trim(rsd("BookNo")) & ","                    ' 1:Booking番号
        sText = sText & Trim(rsd("ContNo")) & ","            ' 2:コンテナNo.
        sWork = Trim(rsd("ShipFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("VslCode")) & ","       ' 3:船名(Full)
        Else
            sText = sText & sWork & ","                      ' 3:船名
        End If
        sText = sText & Trim(rsd("LdVoyage")) & ","          ' 4:次航
        sText = sText & rsd("VslCode") & ","                 ' 5:船名
        sText = sText & rsd("VoyCtrl") & ","                 ' 6:次航(番号)
        sWork = Trim(rsd("ShipName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("VslCode")) & ","       ' 7:船名(略称)
        Else
            sText = sText & sWork & ","                      ' 7:船名(略称)
        End If
        sWork = Trim(rsd("ShipperName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Shipper")) & ","       ' 8:荷主(略称)
        Else
            sText = sText & sWork & ","                      ' 8:荷主(略称)
        End If
        sWork = Trim(rsd("ForName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Forwarder")) & ","     ' 9:海貨(略称)
        Else
            sText = sText & sWork & ","                      ' 9:海貨(略称)
        End If
        sWork = Trim(rsd("TruckName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Trucker")) & ","       '10:陸運(略称)
        Else
            sText = sText & sWork & ","                      '10:陸運(略称)
        End If
        sText = sText & Trim(rsd("ContSize")) & ","          '11:サイズ
        sText = sText & Trim(rsd("ContType")) & ","          '12:タイプ
        sText = sText & Trim(rsd("ContHeight")) & ","        '13:高さ
        sText = sText & Trim(rsd("Remark")) & ","            '14:コメント
        sText = sText & Trim(rsd("ShipCtrl")) & ","          '15:荷主管理番号
        sText = sText & DispDateTime(rsd("WHArTime"),0) & ","     '16:倉庫到着指定日時

        sWork = Trim(rsd("OpeName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("OpeCode")) & ","       '17:港運略称
        Else
            sText = sText & sWork & ","                      '17:港運略称
        End If

'        sText = sText & DispDateTime(rsd("CYRecDate"),0) & ","    '17:CY搬入指定日時
        sText = sText & Trim(rsd("TruckRefFlag")) & ","      '18:陸運参照フラグ
        sText = sText & Trim(rsd("ShipLine")) & ","          '19:船社コード
        sText = sText & Trim(rsd("BookVoyage")) & ","        '20:次航(番号)
        sText = sText & Trim(rsd("PickPlace")) & ","         '21:空コン受取場所
'        sWork = Trim(rsd("ShipperFullName"))
'        If sWork="" Or IsNull(sWork) Then
'            sText = sText & Trim(rsd("Shipper")) & ","       '22:荷主(フル名称)
'        Else
'            sText = sText & sWork & ","                      '22:荷主(フル名称)
'        End If
'        sWork = Trim(rsd("ForFullName"))
'        If sWork="" Or IsNull(sWork) Then
'            sText = sText & Trim(rsd("Forwarder")) & ","     '23:海貨(フル名称)
'        Else
'            sText = sText & sWork & ","                      '23:海貨(フル名称)
'        End If
		
        sText = sText & DispDateTime(rsd("CYCut"),0) & ","    '22:カット日
        sText = sText & Trim(rsd("RecTerminal")) & ","       '23:CY搬入場所

'        sWork = Trim(rsd("TruckFullName"))
'       If sWork="" Or IsNull(sWork) Then
'            sText = sText & Trim(rsd("Trucker")) & ","       '24:陸運(フル名称)
'        Else
'            sText = sText & sWork & ","                      '24:陸運(フル名称)
'        End If

        sText = sText & Trim(rsd("Shipper")) & ","      	' 24:荷主コード

        sText = sText & DispDateTime(rsd("PickDate"),0) & "," '25:空コン搬出指定日
        sText = sText & Trim(rsd("OpeCode")) & ","			'26:港運コード
        sText = sText & Trim(rsd("PickOKFlag")) & ","		'27:空コン引当可否フラグ
        sText = sText & Trim(rsd("PkPlUpdFlag")) & ","		'28:ピック場所変更フラグ
        sText = sText & Trim(rsd("PkDtUpdFlag")) & ","		'29:搬出指定日変更フラグ

        ReDim Preserve Tmp(bWrite)
        Tmp(bWrite) = sText
        bWrite = bWrite + 1

        rsd.MoveNext
    Loop
    rsd.Close

    For i=0 To bWrite-1

        strTmp=Split(Tmp(i),",")

        If strTmp(18)<>"" Then
            sql = "SELECT FullName FROM mShipLine WHERE ShipLine='" & strTmp(18) & "'"
            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                Tmp(i) = Tmp(i) & Trim(rsd("FullName")) & ","       '30:
            Else
                Tmp(i) = Tmp(i) & ","                               '30:
            End If
            rsd.Close
        Else
            Tmp(i) = Tmp(i) & ","                                   '30:
        End If

        If strTmp(5)<>"" Then
            sWhere = "ExportCont.VslCode='" & strTmp(4) & "' AND ExportCont.BookNo='" & strTmp(0) & "' AND ExportCont.ContNo='" & strTmp(1) & "'"
            sql = "SELECT ExportCont.BookNo, ExportCont.ContNo, ExportCont.VslCode, ExportCont.PickPlace, " & _
                  "ExportCont.RFFlag, ExportCont.VoyCtrl, ExportCont.EmpDelTime, ExportCont.WHArTime, " & _
                  "ExportCont.VanTime, ExportCont.ExpArTime, ExportCont.OutTime, ExportCont.SYRecTime, " & _
                  "ExportCont.CYRecPrDate, ExportCont.SYRecFlag, ExportCont.SYPRRecTime, ExportCont.DGFlag, " & _
                  "Container.ContSize, Container.ContType, Container.ContHeight, Container.SealNo, Container.ContWeight, " & _
                  "Container.CargoWeight, Container.ShipLine, " & _
                  "Booking.RecTerminal, Booking.DPort, Booking.WHArSchDate, Booking.RFFlag BRFFlag, Booking.DGFlag BDGFlag, " & _
                  "VslSchedule.LdVoyage, VslSchedule.CYOpen, VslSchedule.CYCut, " & _
                  "mShipLine.FullName ShipLineName, " & _
                  "mVessel.FullName ShipName, " & _
                  "VslPort.ETD, VslPort.TD, VslPort.ETDLong"
            sql = sql & " FROM ExportCont, Container, Booking, VslSchedule, mShipLine, mVessel, VslPort "
            sql = sql & "WHERE (" & sWhere & ") AND Container.VslCode=ExportCont.VslCode AND Container.VoyCtrl=ExportCont.VoyCtrl AND " & _
                  "Container.ContNo=ExportCont.ContNo AND Booking.VslCode=*ExportCont.VslCode AND Booking.VoyCtrl=*ExportCont.VoyCtrl AND " & _
                  "Booking.BookNo=*ExportCont.BookNo AND VslSchedule.VslCode=*ExportCont.VslCode AND VslSchedule.VoyCtrl=*ExportCont.VoyCtrl AND " & _
                  "mShipLine.ShipLine=*Container.ShipLine AND mVessel.VslCode=*ExportCont.VslCode AND " & _
                  "VslPort.VslCode=*ExportCont.VslCode AND VslPort.VoyCtrl=*ExportCont.VoyCtrl AND VslPort.PortCode='" & strPort & "'"

            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
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
                strEmpDelTime = DispDateTime(rsd("EmpDelTime"),0)      '17:空コン受取時刻
                strWHArTime = DispDateTime(rsd("WHArTime"),0)          '18:倉庫到着
                strVanTime = DispDateTime(rsd("VanTime"),0)            '19:バンニング
                strSYRecTime = DispDateTime(rsd("SYRecTime"),0)        '25:ストックヤード搬入時刻
                strExpArTime = DispDateTime(rsd("ExpArTime"),0)        '20:ＣＹ搬入
                strOutTime = DispDateTime(rsd("OutTime"),0)            '21:船積完了
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

                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & DispDateTime(rsd("ETD"),0) & ","        '16:離岸予定時刻(ターミナル)
                End If
                strOut = strOut & DispDateTime(rsd("EmpDelTime"),0) & ","     '17:空コン受取時刻
                strOut = strOut & strWHArTime & ","                           '18:倉庫到着
                strOut = strOut & strVanTime & ","                            '19:バンニング
                strOut = strOut & DispDateTime(rsd("ExpArTime"),0) & ","      '20:ＣＹ搬入
                strOut = strOut & DispDateTime(rsd("OutTime"),0) & ","        '21:船積完了
                If IsNull( strCheck ) Then
                    strOut = strOut & ",,,"
                Else
                    strOut = strOut & DispDateTime(rsd("TD"),0) & ","         '22:離岸完了時刻(ターミナル)
                    strOut = strOut & ",,"
'                    strOut = strOut & DispDateTime(rsd("TA"),0) & ","         '23:着岸完了時刻(仕向港)
'                    strOut = strOut & DispDateTime(rsd("ETA"),0) & ","        '24:着岸予定時刻(仕向港)
                End If
                strOut = strOut & DispDateTime(rsd("SYRecTime"),0) & ","      '25:ストックヤード搬入時刻
                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & DispDateTime(rsd("ETDLong"),0) & ","    '26:離岸スケジュール時刻(ターミナル)
                End If
'                strOut = strOut & DispDateTime(rsd("WHArSchDate"),0) & ","    '27:空倉庫到着スケジュール時刻
                strOut = strOut & ","                                         '27:空倉庫到着スケジュール時刻(通常テーブルには設定されない)

                strOut = strOut & Trim(rsd("CargoWeight")) & ","              '28:貨物重量
                strOut = strOut & Trim(rsd("SYRecFlag")) & ","                '29:ストックヤード搬入予約フラグ
                strOut = strOut & DispDateTime(rsd("SYPRRecTime"),0) & ","    '30:ストックヤード搬入予約日時
                strOut = strOut & DispDateTime(rsd("CYRecPrDate"),0) & ","    '31:ＣＹ搬入指定日時

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
				strOut = strOut & strDGFlag & ","

				strOut = strOut & Trim(rsd("ContType"))              '33:コンテナタイプ

                ' 仕向港データを探す
                strTmp=Split(strOut,",")
                If strTmp(14)<>"" Then
                    rsd.Close

                    sql = "SELECT ExportCont.VoyCtrl, VslPort.TA, VslPort.ETA FROM ExportCont, VslPort " & _
                          "WHERE (" & sWhere & ") AND VslPort.VslCode=ExportCont.VslCode AND " & _
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

                    strWork=""
                    For j=0 To UBound(strTmp)
                        strWork=strWork & strTmp(j) & ","
                    Next
                    strOut=strWork
                End If

            Else
                strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            rsd.Close
        ElseIf strTmp(19)<>"" Then
            sWhere = "Booking.VslCode='" & strTmp(4) & "' AND Booking.BookNo='" & strTmp(0) & "' AND Booking.VoyCtrl=" & strTmp(19)
            sql = "SELECT Booking.BookNo, Booking.RecTerminal, Booking.VslCode, Booking.DPort, Booking.WHArSchDate WHArSchDate, " & _
                  "Booking.RFFlag, Booking.DGFlag, " & _
                  "VslSchedule.LdVoyage, VslSchedule.CYOpen, VslSchedule.CYCut, VslSchedule.ShipLine, " & _
                  "mShipLine.FullName ShipLineName, " & _
                  "mVessel.FullName ShipName, " & _
                  "mPort.FullName DPortName, " & _
                  "A.ETD, A.TD, A.ETDLong, B.TA, B.ETA"
            sql = sql & " FROM Booking, VslSchedule, mShipLine, mVessel, mPort, VslPort A, VslPort B "
            sql = sql & "WHERE (" & sWhere & ") AND VslSchedule.VslCode=Booking.VslCode AND VslSchedule.VoyCtrl=Booking.VoyCtrl AND " & _
                  "mShipLine.ShipLine=*VslSchedule.ShipLine AND mVessel.VslCode=*Booking.VslCode AND " & _
                  "mPort.PortCode=*Booking.DPort AND A.VslCode=*Booking.VslCode AND A.VoyCtrl=*Booking.VoyCtrl AND A.PortCode='" & strPort & "' AND " & _
                  "B.VslCode=*Booking.VslCode AND B.VoyCtrl=*Booking.VoyCtrl AND B.PortCode=*Booking.DPort"

            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' 出力レコードの編集
                strOut = Trim(rsd("BookNo")) & ","                     ' 1:Booking No.
                strOut = strOut & strTmp(1) & ","                      ' 2:コンテナNo.
                strOut = strOut & ","                                  ' 3:空コン受取場所
                strOut = strOut & ","                                  ' 4:サイズ
                strOut = strOut & ","                                  ' 5:高さ
                strOut = strOut & Trim(rsd("RFFlag")) & ","            ' 6:リーファー
                strOut = strOut & Trim(rsd("RecTerminal")) & ","       ' 7:搬入ターミナル名
                strOut = strOut & ","                                  ' 8:シールNo.
                strOut = strOut & ","                                  ' 9:総重量
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
                If IsNull(rsd("DPortName")) Then
                    strOut = strOut & Trim(rsd("DPort")) & ","         '15:仕向港
                Else
                    strOut = strOut & Trim(rsd("DPortName")) & ","     '15:仕向港
                End If

                strOut = strOut & DispDateTime(rsd("ETD"),0) & ","     '16:離岸予定時刻(ターミナル)
                strOut = strOut & ","                                  '17:空コン受取時刻
                strOut = strOut & ","                                  '18:倉庫到着
                strOut = strOut & ","                                  '19:バンニング
                strOut = strOut & ","                                  '20:ＣＹ搬入
                strOut = strOut & ","                                  '21:船積完了
                strOut = strOut & DispDateTime(rsd("TD"),0) & ","      '22:離岸完了時刻(ターミナル)
                strOut = strOut & DispDateTime(rsd("TA"),0) & ","      '23:着岸完了時刻(仕向港)

                strOut = strOut & DispDateTime(rsd("ETA"),0) & ","     '24:着岸予定時刻(仕向港)
                strOut = strOut & ","                                  '25:ストックヤード搬入時刻
                strOut = strOut & DispDateTime(rsd("ETDLong"),0) & "," '26:離岸スケジュール時刻(ターミナル)
                strOut = strOut & DispDateTime(rsd("WHArSchDate"),0) & ","  '27:空倉庫到着スケジュール時刻

                strOut = strOut & ","                                  '28:貨物重量
                strOut = strOut & ","                                  '29:ストックヤード搬入予約フラグ
                strOut = strOut & ","                                  '30:ストックヤード搬入予約日時
                strOut = strOut & ","                                  '31:ＣＹ搬入指定日時
                strOut = strOut & ","                                  '32:危険品
																	   '33:タイプ

            Else
                strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            rsd.Close
        Else
            strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
        End If

        pfp.WriteLine Tmp(i) & strOut & ","
    Next

    SerchMSExpCntnr = bWrite    '該当レコードあったかどうか
End Function

'輸出コンテナCSVファイルタイトル行出力
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]検索モード(ログイン種別)

    If bMode="海貨" Then
        Response.Write "荷主,"
        Response.Write "荷主管理番号,"
    ElseIf bMode="陸運" Then
        Response.Write "海貨,"
    ElseIf bMode="荷主" Then
        Response.Write "荷主管理番号,"
    Else
        Response.Write "荷主,"
        Response.Write "荷主管理番号,"
    End If

    Response.Write "Booking番号,"
    Response.Write "コンテナNo.,"

    If bMode="海貨" Then
        Response.Write "指定陸運業者,"
        Response.Write "倉庫略称,"
    ElseIf bMode="陸運" Then
        Response.Write "荷主,"
        Response.Write "倉庫略称,"
    ElseIf bMode="荷主" Then
        Response.Write "海貨,"
        Response.Write "指定陸運業者,"
    Else
        Response.Write "指定陸運業者,"
        Response.Write "倉庫略称,"
    End If

    Response.Write "空コン受取場所,"
    Response.Write "空コン搬出指定日,"
    Response.Write "サイズ,"
    Response.Write "タイプ,"
    Response.Write "高さ,"
    Response.Write "リーファー,"
    Response.Write "港運コード,"

    Response.Write "シールNo.,"
    Response.Write "貨物重量,"
    Response.Write "総重量,"
    Response.Write "搬入ターミナル名,"
    Response.Write "CYオープン日,"
    Response.Write "CYクローズ日,"

    Response.Write "空コン受取完了時刻,"
    Response.Write "倉庫到着指示時刻,"
    Response.Write "倉庫到着完了時刻,"
    Response.Write "バンニング完了時刻,"
    Response.Write "SY予約時刻,"
    Response.Write "SY搬入完了時刻,"
    Response.Write "CY搬入指示日,"
    Response.Write "CY搬入完了時刻,"
    Response.Write "船積完了時刻,"
    Response.Write "離岸計画時刻,"
    Response.Write "離岸予定時刻,"
    Response.Write "離岸完了時刻,"
    Response.Write "仕向港着岸予定時刻,"
    Response.Write "仕向港着岸完了時刻,"

    Response.Write "船社,"
    Response.Write "船名,"
    Response.Write "Voyage No.,"
    Response.Write "仕向港,"
    Response.Write "危険品"

    Response.Write Chr(13) & Chr(10)
End Sub

'輸出コンテナCSVファイルデータ行出力
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]検索モード(ログイン種別)
    'pfp        [ I ]テンポラリファイル

    ' ダウンロードファイルのレコードがある間繰り返す
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        If bMode="海貨" Then
            ' 荷主名称
            Response.Write anyTmp(21) & ","
            ' 荷主管理番号
            Response.Write anyTmp(14) & ","
        ElseIf bMode="陸運" Then
            ' 海貨
            Response.Write anyTmp(22) & ","
        ElseIf bMode="荷主" Then
            ' 荷主管理番号
            Response.Write anyTmp(14) & ","
        Else
            ' 荷主管理番号
            Response.Write anyTmp(14) & ","
        End If

        ' Booking No.
        Response.Write anyTmp(0) & ","
        ' コンテナNo.
        Response.Write anyTmp(1) & ","

        If bMode="海貨" Then
            ' 指定陸運業者
            Response.Write anyTmp(23) & ","
            ' 倉庫略称
            Response.Write anyTmp(13) & ","
        ElseIf bMode="陸運" Then
            ' 荷主名称
            Response.Write anyTmp(21) & ","
            ' 倉庫略称
            Response.Write anyTmp(13) & ","
        ElseIf bMode="荷主" Then
            ' 海貨
            Response.Write anyTmp(22) & ","
            ' 指定陸運業者
            Response.Write anyTmp(23) & ","
        Else
            ' 指定陸運業者
            Response.Write anyTmp(23) & ","
            ' 倉庫略称
            Response.Write anyTmp(13) & ","
        End If

        ' 空コン受取場所
        If anyTmp(32)<>"" Then
            Response.Write anyTmp(32) & ","
        Else
            Response.Write anyTmp(20) & ","
        End If
        ' 空コン搬出指定日
        Response.Write anyTmp(24) & ","
        ' サイズ
        If anyTmp(33)<>"" Then
            Response.Write anyTmp(33) & ","
        Else
            Response.Write anyTmp(10) & ","
        End If
        ' タイプ
        If anyTmp(62)<>"" Then
            Response.Write anyTmp(62) & ","
        Else
	        Response.Write anyTmp(11) & ","
        End If
        ' 高さ
        If anyTmp(34)<>"" Then
            Response.Write anyTmp(34) & ","
        Else
            Response.Write anyTmp(12) & ","
        End If
        ' リーファー
        If anyTmp(35)="R" Then
            Response.Write "○,"
        ElseIf anyTmp(35)<>"" Then
            Response.Write "－,"
        Else
            If anyTmp(11)<>"" Then
                If anyTmp(11)<>"RF" Then
                    Response.Write "－,"
                Else
                    Response.Write "○,"
                End If
            Else
                Response.Write ","
            End If
        End If
        ' 港運コード
        Response.Write anyTmp(25) & ","
        ' シールNo.
        Response.Write anyTmp(37) & ","
        ' 貨物重量
        If anyTmp(57)<>"" And anyTmp(57)<>"0" Then
            dWeight=anyTmp(57) / 10
            Response.Write dWeight
        Else
            Response.Write "－"
        End If
        Response.Write ","
        ' 総重量
        If anyTmp(38)<>"" And anyTmp(38)<>"0" Then
            dWeight=anyTmp(38) / 10
            Response.Write dWeight
        Else
            Response.Write "－"
        End If
        Response.Write ","
        ' 搬入ターミナル名
        Response.Write anyTmp(36) & ","
        ' ＣＹオープン
        Response.Write anyTmp(39) & ","
        ' ＣＹクローズ
        Response.Write anyTmp(40) & ","

        ' 空コン受取時刻
        Response.Write anyTmp(46) & ","
        ' 倉庫到着スケジュール
        If anyTmp(56)<>"" Then
            Response.Write anyTmp(56) & ","
        Else
            Response.Write anyTmp(15) & ","
        End If
        ' 倉庫到着
        Response.Write anyTmp(47) & ","
        ' バンニング
        Response.Write anyTmp(48) & ","
        ' ストックヤード搬入予約
        sTemp=DispReserveCell(anyTmp(58),anyTmp(59),sColor)
        If anyTmp(58)>="1" And anyTmp(58)<="4" Then
            sTemp=Left(anyTmp(59),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' ストックヤード搬入
        Response.Write anyTmp(54) & ","
        ' ＣＹ搬入指示
        Response.Write Left(anyTmp(60),10) & ","
        ' ＣＹ搬入
        Response.Write anyTmp(49) & ","
        ' 船積完了
        Response.Write anyTmp(50) & ","
        ' 離岸スケジュール
        Response.Write anyTmp(25) & ","
        ' 離岸予定時刻
        Response.Write anyTmp(45) & ","
        ' 離岸完了時刻
        Response.Write anyTmp(51) & ","
        ' 着岸予定時刻
        Response.Write anyTmp(53) & ","
        ' 着岸完了時刻
        Response.Write anyTmp(52) & ","

        ' 船社名
        If anyTmp(41)<>"" Then
            Response.Write anyTmp(41) & ","
        Else
            Response.Write anyTmp(29) & ","
        End If
        ' 船名
        If anyTmp(32)<>"" Then
            Response.Write anyTmp(32) & ","
        Else
            Response.Write anyTmp(2) & ","
        End If
        ' 次航
        If anyTmp(43)<>"" Then
            Response.Write anyTmp(43) & ","
        Else
            Response.Write anyTmp(3) & ","
        End If
        ' 仕向港
        Response.Write anyTmp(44) & ","

        ' 危険品
        If anyTmp(61)="H" Then
            Response.Write "○"
        ElseIf anyTmp(61)<>"" Then
            Response.Write "－"
        Else
            Response.Write ""
        End If

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>