<%
'輸入照会関係関数

'輸入コンテナ照会
Function SerchMSImpCntnr(conn, rsd, pfp, sWhere, sOrder)
    'conn       [ I ]DB接続Object
    'rsd        [ I ]レコードセットObject
    'pfp        [ I ]テンポラリファイル
    'sWhere     [ I ]検索条件
    'sOrder     [ I ]ソート条件

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    bWrite = 0        '出力レコード件数

    sql = "SELECT ImportCargoInfo.VslCode, ImportCargoInfo.DsVoyage, ImportCargoInfo.ContNo, " & _
          "ImportCargoInfo.BLNo, ImportCargoInfo.Shipper, ImportCargoInfo.Forwarder, ImportCargoInfo.Trucker, " & _
          "ImportCargoInfo.WHArTime, ImportCargoInfo.ContSize, ImportCargoInfo.ContType, ImportCargoInfo.Remark, " & _
          "ImportCargoInfo.TruckRefFlag, " & _
          "VslSchedule.ShipLine, " & _
          "ImportCont.VoyCtrl, " & _
          "BL.VoyCtrl BLVoyage, " & _
          "mVessel.FullName ShipFullName, mVessel.NameAbrev ShipName, " & _
          "mShipper.NameAbrev ShipperName, mShipper.FullName ShipperFullName, " & _
          "mForwarder.NameAbrev ForName, mForwarder.FullName ForFullName, " & _
          "mTrucker.NameAbrev TruckName, mTrucker.FullName TruckFullName"
    sql = sql & " FROM ImportCargoInfo, VslSchedule, ImportCont, BL, mVessel, mShipper, mForwarder, mTrucker "
    sql = sql & "WHERE (" & sWhere & ") AND VslSchedule.VslCode=*ImportCargoInfo.VslCode AND VslSchedule.DsVoyage=*ImportCargoInfo.DsVoyage AND " & _
          "ImportCont.VslCode=*ImportCargoInfo.VslCode AND ImportCont.BLNo=*ImportCargoInfo.BLNo AND ImportCont.ContNo=*ImportCargoInfo.ContNo AND " & _
          "BL.VslCode=*ImportCargoInfo.VslCode AND BL.BLNo=*ImportCargoInfo.BLNo AND " & _
          "mVessel.VslCode=*ImportCargoInfo.VslCode AND " & _
          "mShipper.Shipper=*ImportCargoInfo.Shipper AND " & _
          "mForwarder.Forwarder=*ImportCargoInfo.Forwarder AND " & _
          "mTrucker.Trucked=*ImportCargoInfo.Trucker"
    sql = sql & " ORDER BY " & sOrder

    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1

    Do While Not rsd.EOF
        sText=""

        ' 出力レコードの編集
        sText = Trim(rsd("BLNo")) & ","                      ' 1:BL番号
        sText = sText & Trim(rsd("ContNo")) & ","            ' 2:コンテナNo.
        sWork = Trim(rsd("ShipFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("VslCode")) & ","       ' 3:船名(Full)
        Else
            sText = sText & sWork & ","                      ' 3:船名
        End If
        sText = sText & Trim(rsd("DsVoyage")) & ","          ' 4:次航
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
        sText = sText & Trim(rsd("Remark")) & ","            '13:コメント
        sText = sText & DispDateTime(rsd("WHArTime"),0) & ","     '14:倉庫到着指定日時
        sText = sText & Trim(rsd("TruckRefFlag")) & ","      '15:陸運参照フラグ
        sText = sText & Trim(rsd("ShipLine")) & ","          '16:船社コード
        sText = sText & Trim(rsd("BLVoyage")) & ","          '17:次航(番号)
        sWork = Trim(rsd("ShipperFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Shipper")) & ","       '18:荷主(フル名称)
        Else
            sText = sText & sWork & ","                      '18:荷主(フル名称)
        End If
        sWork = Trim(rsd("ForFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Forwarder")) & ","     '19:海貨(フル名称)
        Else
            sText = sText & sWork & ","                      '19:海貨(フル名称)
        End If
        sWork = Trim(rsd("TruckFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Trucker")) & ","       '20:陸運(フル名称)
        Else
            sText = sText & sWork & ","                      '20:陸運(フル名称)
        End If

        ReDim Preserve Tmp(bWrite)
        Tmp(bWrite) = sText
        bWrite = bWrite + 1

        rsd.MoveNext
    Loop
    rsd.Close

    For i=0 To bWrite-1

        strTmp=Split(Tmp(i),",")

        If strTmp(15)<>"" Then
            sql = "SELECT NameAbrev, FullName FROM mShipLine WHERE ShipLine='" & strTmp(15) & "'"
            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                Tmp(i) = Tmp(i) & Trim(rsd("NameAbrev")) & "," & Trim(rsd("FullName")) & ",,,,,,,,,"       '21: 22:
            Else
                Tmp(i) = Tmp(i) & ",,,,,,,,,,"                  '21: 22:
            End If
            rsd.Close
        Else
            Tmp(i) = Tmp(i) & ",,,,,,,,,,"                      '21: 22:
        End If

        If strTmp(5)<>"" Then
            sWhere = "ImportCont.VslCode='" & strTmp(4) & "' AND ImportCont.BLNo='" & strTmp(0) & "' AND ImportCont.ContNo='" & strTmp(1) & "'"
            sql = "SELECT ImportCont.BLNo, ImportCont.ContNo, ImportCont.RFFlag, ImportCont.DGFlag, " & _
                  "ImportCont.VslCode, ImportCont.VoyCtrl, ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, " & _
                  "ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, ImportCont.ReturnPlace, " & _
                  "ImportCont.InTime, ImportCont.CYDelTime, ImportCont.SYDelTime, ImportCont.WHArDate, ImportCont.DevanTime, " & _
                  "ImportCont.ReturnTime, ImportCont.CustStatus, ImportCont.FreeTimeExt, ImportCont.WHArSchDate ImCntnrWHArSch, " & _
                  "ImportCont.SYDelFlag, ImportCont.SYPRDelTime, " & _
                  "Container.ContSize, Container.ContHeight, Container.ContWeight, Container.ListDate, Container.StackPlace, Container.ShipLine, " & _
                  "InspectCont.BotStatus, InspectCont.ZooStatus, BL.RecTerminal, BL.IndvRecFlag, BL.IndvRecSchTime, BL.WHArSchDate BLWHArSch, BL.LPort, " & _
                  "BL.RFFlag BRFFlag, BL.DGFlag BDGFlag, " & _
                  "VslSchedule.DsVoyage, VslSchedule.BatchRecSchTime, " & _
                  "mShipLine.FullName ShipLineName, " & _
                  "mVessel.FullName ShipName, " & _
                  "VslPort.ETA, VslPort.TA, VslPort.ETALong"
            sql = sql & " FROM ImportCont, Container, InspectCont, BL, VslSchedule, mShipLine, mVessel, VslPort "
            sql = sql & "WHERE (" & sWhere & ") AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl AND " & _
                  "Container.ContNo=ImportCont.ContNo AND InspectCont.VslCode=*ImportCont.VslCode AND InspectCont.VoyCtrl=*ImportCont.VoyCtrl AND " & _
                  "InspectCont.ContNo=*ImportCont.ContNo AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl AND " & _
                  "BL.BLNo=*ImportCont.BLNo AND VslSchedule.VslCode=*ImportCont.VslCode AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl AND " & _
                  "mShipLine.ShipLine=*Container.ShipLine AND mVessel.VslCode=*ImportCont.VslCode AND " & _
                  "VslPort.VslCode=*ImportCont.VslCode AND VslPort.VoyCtrl=*ImportCont.VoyCtrl AND VslPort.PortCode='" & strPort & "'"
            sql = sql & " ORDER BY ImportCont.BLNo, ImportCont.ContNo"

            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' 通関／保税輸送のチェック
                If Trim(rsd("OLTICFlag"))="I" Then
                   If Trim(rsd("OLTICNo"))<>"" Then
                        strTsukan = "I"
                    Else
                        strTsukan = ""
                    End If
                Else
                   If Now>=rsd("OLTDateFrom") And Now<=rsd("OLTDateTo") Then
                        strTsukan = "O"
                    Else
                        strTsukan = ""
                    End If
                End If
                If DispDateTime(rsd("CYDelTime"),0)<>"" Then           ' 搬出されていたら○とする
                    If IsNull(rsd("OLTDateFrom")) Or IsNull(rsd("OLTDateTo")) Then
                        strTsukan = "S"
                    Else
                        strTsukan = "T"
                    End If
                End If

                ' 出力レコードの編集
                strOut = Trim(rsd("BLNo")) & ","                       ' 1:BL番号
                strOut = strOut & Trim(rsd("ContNo")) & ","            ' 2:コンテナNo.

                strCheck=rsd("ShipName")                '船名がなかったら、離岸・着岸時刻を表示しないようにする

                If IsNull( strCheck ) Then
                    strOut = strOut & ",,"
                Else
                    strOut = strOut & DispDateTime(rsd("ETA"),0) & "," ' 3:着岸予定時刻
                    strOut = strOut & DispDateTime(rsd("TA"),0) & ","  ' 4:着岸完了時刻
                End If
                ' 5:搬出可否
                If DispDateTime(rsd("CYDelTime"),0)<>"" Then
                    strOut = strOut & "S,"
                ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
                   (rsd("FreeTime")>=Now Or rsd("FreeTimeExt")>=Now) And _
                   rsd("DelPermitDate")<=Now Then
                    strOut = strOut & "Y,"
                Else
                    strOut = strOut & ","
                End If
'                strOut = strOut & Trim(rsd("StackPlace")) & ","        ' 6:搬出ターミナル名
                strOut = strOut & Trim(rsd("RecTerminal")) & ","       ' 6:搬出ターミナル名

                If IsNull(rsd("ShipLineName")) Then
                    strOut = strOut & Trim(rsd("ShipLine")) & ","      ' 7:船社名
                Else
                    strOut = strOut & Trim(rsd("ShipLineName")) & ","  ' 7:船社名
                End If
                If IsNull(rsd("ShipName")) Then
                    strOut = strOut & Trim(rsd("VslCode")) & ","       ' 8:船名
                Else
                    strOut = strOut & Trim(rsd("ShipName")) & ","      ' 8:船名
                End If
                strOut = strOut & Trim(rsd("DsVoyage")) & ","          ' 9:次航
                strOut = strOut & Trim(rsd("LPort")) & ","             '10:仕出港(コード)
                strOut = strOut & Trim(rsd("ReturnPlace")) & ","       '11:返却場所

                ' 時系列のチェック
                strCYDelTime = DispDateTime(rsd("CYDelTime"),0)        '14:ヤード搬出
                strSYDelTime = DispDateTime(rsd("SYDelTime"),0)        '31:ストックヤード搬出時刻
                strWHArDate = DispDateTime(rsd("WHArDate"),0)          '15:倉庫到着
                strDevanTime = DispDateTime(rsd("DevanTime"),0)        '16:デバン
                strReturnTime = DispDateTime(rsd("ReturnTime"),0)      '17:空コン返却
                strOld = strSYDelTime
                If strOld = "" Then
                    strOld = strCYDelTime
                End If
                If strWHArDate<strOld Then
                    strWHArDate = ""                      ' 倉庫到着がKACCS,Hits時刻より過去のとき表示しない
                End If
                If strReturnTime<>"" And strWHArDate>strReturnTime Then
                    strWHArDate = ""                      ' 倉庫到着が空コン返却時刻より未来のとき表示しない
                End If
                If strWHArDate<>"" Then
                    strOld = strWHArDate
                End If
                If strDevanTime<strOld Then
                    strDevanTime = ""                     ' デバンがKACCS,Hits時刻より過去のとき表示しない
                End If
                If strReturnTime<>"" And strDevanTime>strReturnTime Then
                    strDevanTime = ""                     ' デバンが空コン返却時刻より未来のとき表示しない
                End If

                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & ","
'                    strOut = strOut & DispDateTime(rsd("TD"),0) & ","     '12:離岸完了時刻
                End If
                strOut = strOut & DispDateTime(rsd("InTime"),0) & ","     '13:ヤード搬入
                strOut = strOut & DispDateTime(rsd("CYDelTime"),0) & ","  '14:ヤード搬出
                strOut = strOut & strWHArDate & ","                       '15:倉庫到着
                strOut = strOut & strDevanTime & ","                      '16:デバン
                strOut = strOut & DispDateTime(rsd("ReturnTime"),0) & "," '17:空コン返却

                '18:動植物
                strBotStat=Trim(rsd("BotStatus"))
                strZooStat=Trim(rsd("ZooStatus"))
                If strBotStat="S" Or strBotStat="C" Then
                    strOut = strOut & strBotStat & ","
                ElseIf strZooStat="S" Or strZooStat="C" Then
                    strOut = strOut & strZooStat & ","
                Else
                    strOut = strOut & ","
                End If
                strOut = strOut & DispDateTime(rsd("ListDate"),0) & ","    '19:搬入確認完了時刻
                strOut = strOut & strTsukan & ","                          '20:通関／保税輸送
                strOut = strOut & Trim(rsd("CustStatus")) & ","            '21:通関審査状況
                strOut = strOut & Trim(rsd("DOStatus")) & ","              '22:商取引DO発行
                '23:フリータイム
                strFree = DispDateTime(rsd("FreeTime"),10)
                strFreeExt = DispDateTime(rsd("FreeTimeExt"),10)
                If strFreeExt<>"" Then
                    strFree = strFreeExt
                End If
                strOut = strOut & strFree & ","

                strOut = strOut & Trim(rsd("ContSize")) & ","          '24:サイズ
                strOut = strOut & Trim(rsd("ContHeight")) & ","        '25:高さ
                '26:リーファー
		        strRFFlag = Trim(rsd("RFFlag"))
		        If IsNull(strRFFlag) Or strRFFlag="" Then
		            strRFFlag = Trim(rsd("BRFFlag"))
		        End If
		        If IsNull(strRFFlag) Or strRFFlag="" Then
		            If Not IsNull(rsd("ContSize")) Then
						strRFFlag = "-"
		            End If
		        End If
				strOut = strOut & strRFFlag & ","
                strOut = strOut & Trim(rsd("ContWeight")) & ","        '27:重量
                '28:危険物
		        strDGFlag = Trim(rsd("DGFlag"))
		        If IsNull(strDGFlag) Or strDGFlag="" Then
		            strDGFlag = Trim(rsd("BDGFlag"))
		        End If
		        If IsNull(strDGFlag) Or strDGFlag="" Then
		            If Not IsNull(rsd("ContSize")) Then
						strDGFlag = "-"
		            End If
		        End If
				strOut = strOut & strDGFlag & ","

                strOut = strOut & DispDateTime(rsd("OLTDateFrom"),10) & "," '29:保税輸送(From)
                strOut = strOut & DispDateTime(rsd("OLTDateTo"),10) & ","   '30:保税輸送(To)

                strOut = strOut & DispDateTime(rsd("SYDelTime"),0) & ","         '31:ストックヤード搬出時刻
                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & DispDateTime(rsd("ETALong"),0) & ","       '32:着岸スケジュール
                End If
                '33:搬入確認予定時刻
                strIndv = DispDateTime(rsd("IndvRecSchTime"),0)
                If strIndv="" Then
                    strIndv = DispDateTime(rsd("BatchRecSchTime"),0)
                End If
                strOut = strOut & strIndv & ","
                strOut = strOut & Trim(rsd("IndvRecFlag")) & ","                 '34:個別搬入
                '35:倉庫到着スケジュール
                strWHA = DispDateTime(rsd("ImCntnrWHArSch"),0)
                If strWHA="" Then
                    strWHA = DispDateTime(rsd("BLWHArSch"),0)
                End If
                strOut = strOut & strWHA & ","
                strOut = strOut & Trim(rsd("SYDelFlag")) & ","                   '36:ストックヤード搬出予約フラグ
                strWork=DispDateTime(rsd("SYPRDelTime"),0)
                If Len(strWork)>13 Then
                    strTemp=Right("0" & CInt(Mid(strWork, 12, 2))+1, 2)
                    strWork=Left(strWork, 11) & strTemp & Right(strWork, 3)
                End If
                strOut = strOut & strWork & ","                                  '37:ストックヤード搬出予約時刻

                rsd.Close

                ' 積港データを探す
                strTmp=Split(strOut,",")
                If strTmp(9)<>"" Then
                    sql = "SELECT ImportCont.VoyCtrl, VslPort.TD FROM ImportCont, VslPort " & _
                          "WHERE (" & sWhere & ") AND VslPort.VslCode=ImportCont.VslCode AND " & _
                          "VslPort.VoyCtrl=ImportCont.VoyCtrl AND VslPort.PortCode='" & strTmp(9) & "'"
                    'SQLを発行して検索
                    rsd.Open sql, conn, 0, 1, 1

                    If Not rsd.EOF Then
                        If strTmp(7)<>"" Then    '船名があったら
                            strTmp(11) = DispDateTime(rsd("TD"),0)     '12:離岸完了時刻
                        End If
                    End If
                    rsd.Close

                    sql = "SELECT mPort.FullName From mPort " & _
                          "WHERE mPort.PortCode='" & strTmp(9) & "'"
                    'SQLを発行して検索
                    rsd.Open sql, conn, 0, 1, 1

                    If Not rsd.EOF Then
                        If Not IsNull(rsd("FullName")) Then
                            strTmp(9) = Trim(rsd("FullName"))          '10:仕出港
                        End If
                    End If
                    rsd.Close

                    strWork=""
                    For j=0 To UBound(strTmp)-1
                        strWork=strWork & strTmp(j) & ","
                    Next
                    strOut=strWork
                End If

                ' 前港データを探す
                sql = "SELECT ImportCont.VoyCtrl, B.TD, mPort.FullName FROM ImportCont, VslPort A, VslPort B, mPort " & _
                      "WHERE (" & sWhere & ") AND A.VslCode=ImportCont.VslCode AND " & _
                      "A.VoyCtrl=ImportCont.VoyCtrl AND A.PortCode='" & strPort & "' AND " & _
                      "B.VslCode=ImportCont.VslCode AND B.VoyCtrl=ImportCont.VoyCtrl AND B.CallSeq=(A.CallSeq-1) AND " & _
                      "mPort.PortCode=*B.PortCode"
                'SQLを発行して検索
                rsd.Open sql, conn, 0, 1, 1

                If Not rsd.EOF Then
                    If IsNull( strCheck ) Then
                        strOut = strOut & ","
                    Else
                        strOut = strOut & DispDateTime(rsd("TD"),0) & ","      '38:前港離岸完了
                    End If
                    strOut = strOut & Trim(rsd("FullName"))                    '39:前港
                Else
                    strOut = strOut & ","
                End If

            Else
                strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            rsd.Close
        ElseIf  strTmp(16)<>"" Then
            sWhere = "BL.VslCode='" & strTmp(4) & "' AND BL.BLNo='" & strTmp(0) & "' AND BL.VoyCtrl=" & strTmp(16)
            sql = "SELECT BL.BLNo, BL.RecTerminal, BL.IndvRecFlag, BL.IndvRecSchTime, BL.WHArSchDate BLWHArSch, " & _
                  "BL.RFFlag, BL.DGFlag, " & _
                  "VslSchedule.DsVoyage, " & _
                  "mShipLine.FullName ShipLineName, " & _
                  "mVessel.FullName ShipName, " & _
                  "PA.FullName LPortName, " & _
                  "A.ETA, A.TA, A.ETALong, B.TD, C.TD BeforeTD, " & _
                  "PB.FullName BeforePort"
            sql = sql & " FROM BL, VslSchedule, mShipLine, mVessel, mPort PA, mPort PB, VslPort A, VslPort B,  VslPort C "
            sql = sql & "WHERE (" & sWhere & ") AND VslSchedule.VslCode=BL.VslCode AND VslSchedule.VoyCtrl=BL.VoyCtrl AND " & _
                  "mShipLine.ShipLine=*VslSchedule.ShipLine AND mVessel.VslCode=*BL.VslCode AND " & _
                  "PA.PortCode=*BL.LPort AND A.VslCode=BL.VslCode AND A.VoyCtrl=BL.VoyCtrl AND A.PortCode='" & strPort & "' AND " & _
                  "B.VslCode=*BL.VslCode AND B.VoyCtrl=*BL.VoyCtrl AND B.PortCode=*BL.LPort AND " & _
                  "C.VslCode=BL.VslCode AND C.VoyCtrl=BL.VoyCtrl AND C.CallSeq=(A.CallSeq-1) AND " & _
                  "PB.PortCode=*C.PortCode"

            'SQLを発行して検索
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' 出力レコードの編集
                strOut = Trim(rsd("BLNo")) & ","                       ' 1:BL番号
                strOut = strOut & strTmp(1) & ","                      ' 2:コンテナNo.
                strOut = strOut & DispDateTime(rsd("ETA"),0) & ","     ' 3:着岸予定時刻
                strOut = strOut & DispDateTime(rsd("TA"),0) & ","      ' 4:着岸完了時刻
                strOut = strOut & ","                                  ' 5:搬出可否
                strOut = strOut &  Trim(rsd("RecTerminal")) & ","      ' 6:搬出ターミナル名

                strOut = strOut & Trim(rsd("ShipLineName")) & ","      ' 7:船社名
                strOut = strOut & Trim(rsd("ShipName")) & ","          ' 8:船名
                strOut = strOut & Trim(rsd("DsVoyage")) & ","          ' 9:次航
                strOut = strOut & Trim(rsd("LPortName")) & ","         '10:仕出港
                strOut = strOut & ","                                  '11:返却場所

                strOut = strOut & DispDateTime(rsd("TD"),0) & ","      '12:離岸完了時刻
                strOut = strOut & ","                                  '13:ヤード搬入
                strOut = strOut & ","                                  '14:ヤード搬出
                strOut = strOut & ","                                  '15:倉庫到着
                strOut = strOut & ","                                  '16:デバン
                strOut = strOut & ","                                  '17:空コン返却

                strOut = strOut & ","                                  '18:動植物
                strOut = strOut & ","                                  '19:搬入確認完了時刻
                strOut = strOut & ","                                  '20:通関／保税輸送
                strOut = strOut & ","                                  '21:通関審査状況
                strOut = strOut & ","                                  '22:商取引DO発行
                strOut = strOut & ","                                  '23:フリータイム

                strOut = strOut & ","                                  '24:サイズ
                strOut = strOut & ","                                  '25:高さ
                strOut = strOut & Trim(rsd("RFFlag")) & ","            '26:リーファー
                strOut = strOut & ","                                  '27:重量
                strOut = strOut & Trim(rsd("DGFlag")) & ","            '28:危険物

                strOut = strOut & ","                                  '29:保税輸送(From)
                strOut = strOut & ","                                  '30:保税輸送(To)

                strOut = strOut & ","                                  '31:ストックヤード搬出時刻
                strOut = strOut & DispDateTime(rsd("ETALong"),0) & ","           '32:着岸スケジュール
                strOut = strOut & DispDateTime(rsd("IndvRecSchTime"),0) & ","    '33:搬入確認予定時刻
                strOut = strOut & Trim(rsd("IndvRecFlag")) & ","                 '34:個別搬入
                strOut = strOut & DispDateTime(rsd("BLWHArSch"),0) & ","         '35:倉庫到着スケジュール
                strOut = strOut & ","                                  '36:ストックヤード搬出予約フラグ
                strOut = strOut & ","                                  '37:ストックヤード搬出予約時刻
                strOut = strOut & DispDateTime(rsd("BeforeTD"),0) & ","          '38:前港離岸完了
                strOut = strOut & Trim(rsd("BeforePort"))                        '39:前港
            Else
                strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            rsd.Close
        Else
            strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
        End If

        pfp.WriteLine Tmp(i) & strOut & ","
    Next

    SerchMSImpCntnr = bWrite    '該当レコードあったかどうか
End Function

'輸入コンテナCSVファイルタイトル行出力
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]検索モード(ログイン種別)

    If bMode="海貨" Then
        Response.Write "船名,"
        Response.Write "Voyage No.,"
        Response.Write "荷主,"
        Response.Write "船社,"
    ElseIf bMode="陸運" Then
        Response.Write "海貨,"
    Else
        Response.Write "船名,"
        Response.Write "Voyage No.,"
        Response.Write "海貨,"
        Response.Write "船社,"
    End If

    Response.Write "BL番号,"
    Response.Write "コンテナNo.,"

    If bMode="海貨" Then
        Response.Write "指定陸運業者,"
    ElseIf bMode="陸運" Then
        Response.Write "倉庫略称,"
    Else
        Response.Write "指定陸運業者,"
    End If

    Response.Write "仕出港離岸完了時刻,"
    Response.Write "前港離岸完了時刻,"
    Response.Write "CY着岸計画,"
    Response.Write "CY着岸予定時刻,"
    Response.Write "CY着岸完了時刻,"
    Response.Write "CY搬入完了時刻,"
    Response.Write "CY搬出完了時刻,"
    Response.Write "SY予約時刻,"
    Response.Write "SY搬出完了時刻,"
    Response.Write "倉庫到着指示時刻,"
    Response.Write "倉庫到着完了時刻,"
    Response.Write "デバン完了時刻,"
    Response.Write "空コン返却時刻,"

    Response.Write "搬入確認予定時刻,"
    Response.Write "搬入確認完了時刻,"
    Response.Write "動植物検疫,"
    Response.Write "個別搬入,"
    Response.Write "通関/保税輸送,"
    Response.Write "DO発行,"
    Response.Write "フリータイム,"
    Response.Write "搬出可否,"

    Response.Write "サイズ,"
    Response.Write "高さ,"
    Response.Write "リーファー,"
    Response.Write "総重量,"
    Response.Write "危険物,"
    Response.Write "搬出ターミナル名,"
    Response.Write "ストックヤード利用,"
    Response.Write "返却先,"

    If bMode="陸運" Then
        Response.Write "船社,"
        Response.Write "船名,"
        Response.Write "Voyage No.,"
    End If
    Response.Write "仕出港,"
    Response.Write "前港"

'    Response.Write "保税輸送(From),"
'    Response.Write "保税輸送(To)"

    Response.Write Chr(13) & Chr(10)
End Sub

'輸入コンテナCSVファイルデータ行出力
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]検索モード(ログイン種別)
    'pfp        [ I ]テンポラリファイル

    ' ダウンロードファイルのレコードがある間繰り返す
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        If bMode="海貨" Then
            ' 船名
            If anyTmp(37)<>"" Then
                Response.Write anyTmp(37) & ","
            ElseIf anyTmp(2)<>"" Then
                Response.Write anyTmp(2) & ","
            Else
                Response.Write ","
            End If
            ' Voyage No.
            If anyTmp(38)<>"" Then
                Response.Write anyTmp(38) & ","
            ElseIf anyTmp(3)<>"" Then
                Response.Write anyTmp(3) & ","
            Else
                Response.Write ","
            End If
            ' 荷主
            Response.Write anyTmp(17) & ","
            ' 船社
            If anyTmp(36)<>"" Then
                Response.Write anyTmp(36) & ","
            ElseIf anyTmp(21)<>"" Then
                Response.Write anyTmp(21) & ","
            Else
                Response.Write anyTmp(15) & ","
            End If
        ElseIf bMode="陸運" Then
            ' 海貨
            Response.Write anyTmp(18) & ","
        Else
            ' 船名
            If anyTmp(37)<>"" Then
                Response.Write anyTmp(37) & ","
            ElseIf anyTmp(2)<>"" Then
                Response.Write anyTmp(2) & ","
            Else
                Response.Write ","
            End If
            ' Voyage No.
            If anyTmp(38)<>"" Then
                Response.Write anyTmp(38) & ","
            ElseIf anyTmp(3)<>"" Then
                Response.Write anyTmp(3) & ","
            Else
                Response.Write ","
            End If
            ' 海貨
            Response.Write anyTmp(18) & ","
            ' 船社
            If anyTmp(36)<>"" Then
                Response.Write anyTmp(36) & ","
            ElseIf anyTmp(21)<>"" Then
                Response.Write anyTmp(21) & ","
            Else
                Response.Write anyTmp(15) & ","
            End If
        End If

        ' BL番号
        Response.Write anyTmp(0) & ","
        ' コンテナNo.
        Response.Write anyTmp(1) & ","

        If bMode="海貨" Then
            ' 指定陸運業者
            Response.Write anyTmp(19) & ","
        ElseIf bMode="陸運" Then
            ' 倉庫略称
            Response.Write anyTmp(12) & ","
        Else
            ' 指定陸運業者
            Response.Write anyTmp(19) & ","
        End If

        ' 離岸完了時刻(仕出港)
        Response.Write anyTmp(41) & ","
        ' 離岸完了時刻(前港)
        Response.Write anyTmp(67) & ","
        ' 着岸スケジュール
        Response.Write Left(anyTmp(61),10) & ","
        ' 着岸予定時刻
        Response.Write anyTmp(32) & ","
        ' 着岸完了時刻
        Response.Write anyTmp(33) & ","
        ' ヤード搬入
        Response.Write anyTmp(42) & ","
        ' ヤード搬出
        Response.Write anyTmp(43) & ","
        ' ストックヤード搬出予約
        sTemp=DispReserveCell(anyTmp(65),anyTmp(66),sColor)
        If anyTmp(65)>="1" And anyTmp(65)<="4" Then
            sTemp=Left(anyTmp(66),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' ストックヤード搬出
        Response.Write anyTmp(60) & ","
        ' 倉庫到着スケジュール
        If anyTmp(64)<>"" Then
            Response.Write anyTmp(64) & ","
        Else
            Response.Write anyTmp(13) & ","
        End If
        ' 倉庫到着完了時刻
        Response.Write anyTmp(44) & ","
        ' デバン
        Response.Write anyTmp(45) & ","
        ' 空コン返却
        Response.Write anyTmp(46) & ","

        ' 搬入確認予定時刻
        Response.Write anyTmp(62) & ","
        ' 搬入確認完了時刻
        Response.Write Left(anyTmp(48),10) & ","
        ' 動植物
        If anyTmp(47)="S" Then
            Response.Write "×,"
        ElseIf anyTmp(47)="C" Then
            Response.Write "○,"
        Else
            Response.Write "－,"
        End If
        ' 個別搬入
        If anyTmp(63)<>"" Then
            Response.Write "○,"
        Else
            Response.Write "×,"
        End If
        ' 通関／保税輸送
        If anyTmp(49)<>"" Then
            Response.Write "○,"
        Else
            Response.Write "×,"
        End If
        ' 商取引DO発行
        If anyTmp(51)<>"Y" Then
            Response.Write "×,"
        Else
            Response.Write "○,"
        End If
        ' フリータイム
        Response.Write Left(anyTmp(52),10) & ","
        ' 搬出可否
        If anyTmp(34)="Y" Then
            Response.Write "○,"
        ElseIf anyTmp(34)="S" Then
            Response.Write "済,"
        Else
            Response.Write "×,"
        End If

        ' サイズ
        If anyTmp(53)<>"" Then
            Response.Write anyTmp(53) & ","
        Else
            Response.Write anyTmp(10) & ","
        End if
        ' 高さ
        Response.Write anyTmp(54) & ","
        ' リーファー
        If anyTmp(55)="R" Then
            Response.Write "○,"
        ElseIf anyTmp(55)<>"" Then
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
        ' 重量
        If anyTmp(56)<>"" And anyTmp(56)<>"0" Then
            dWeight=anyTmp(56) / 10
            Response.Write dWeight
        Else
            Response.Write "－"
        End If
        Response.Write ","
        ' 危険物
        If anyTmp(57)="H" Then
            Response.Write "○,"
        ElseIf anyTmp(57)<>"" Then
            Response.Write "－,"
        Else
            Response.Write ","
        End If
        ' 搬出ターミナル名
        Response.Write anyTmp(35) & ","
        ' ストックヤード利用
        If anyTmp(65)>="1" And anyTmp(65)<="4" Then
            Response.Write "○,"
        Else
            Response.Write "×,"
        End If
        ' 返却場所
        Response.Write anyTmp(40) & ","

        If bMode="陸運" Then
            ' 船社名
            If anyTmp(36)<>"" Then
                Response.Write anyTmp(36) & ","
            ElseIf anyTmp(21)<>"" Then
                Response.Write anyTmp(21) & ","
            Else
                Response.Write anyTmp(15) & ","
            End If
            ' 船名
            If anyTmp(37)<>"" Then
                Response.Write anyTmp(37) & ","
            Else
                Response.Write anyTmp(2) & ","
            End If
            ' 次航
            If anyTmp(38)<>"" Then
                Response.Write anyTmp(38) & ","
            Else
                Response.Write anyTmp(3) & ","
            End If
        End If
        ' 仕出港
        Response.Write anyTmp(39) & ","
        ' 前港
        Response.Write anyTmp(68)

        ' 保税輸送(From)
'        Response.Write anyTmp(58) & ","
        ' 保税輸送(To)
'        Response.Write anyTmp(59)

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>