<%
'輸入照会関係関数

'輸入コンテナ照会
Function SerchImpCntnr(conn, rsd, pfp, sWhere)
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

    sql = "SELECT ImportCont.BLNo, ImportCont.ContNo, ImportCont.RFFlag, ImportCont.DGFlag, " & _
          "ImportCont.VslCode, ImportCont.VoyCtrl, ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, " & _
          "ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, ImportCont.ReturnPlace, " & _
          "ImportCont.InTime, ImportCont.CYDelTime, ImportCont.SYDelTime, ImportCont.WHArDate, ImportCont.DevanTime, " & _
          "ImportCont.ReturnTime, ImportCont.CustStatus, ImportCont.FreeTimeExt, ImportCont.WHArSchDate ImCntnrWHArSch, " & _
          "ImportCont.SYDelFlag, ImportCont.SYPRDelTime, " & _
          "Container.ContSize, Container.ContHeight, Container.ContWeight, Container.ListDate, Container.StackPlace, Container.ShipLine, " & _
          "InspectCont.BotStatus, InspectCont.ZooStatus, BL.RecTerminal, BL.IndvRecFlag, BL.IndvRecSchTime, BL.WHArSchDate BLWHArSch, BL.LPort, " & _
          "BL.RFFlag BRFFlag, BL.DGFlag BDGFlag, " & _
          "ImportCargoInfo.WHArTime, VslSchedule.DsVoyage, VslSchedule.BatchRecSchTime, " & _
          "mShipLine.FullName ShipLineName, " & _
          "mVessel.FullName ShipName, " & _
          "VslPort.ETA, VslPort.TA, VslPort.ETALong"
    sql = sql & " FROM ImportCont, Container, InspectCont, BL, ImportCargoInfo, VslSchedule, mShipLine, mVessel, VslPort "
    sql = sql & "WHERE (" & sWhere & ") AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl AND " & _
          "Container.ContNo=ImportCont.ContNo AND InspectCont.VslCode=*ImportCont.VslCode AND InspectCont.VoyCtrl=*ImportCont.VoyCtrl AND " & _
          "InspectCont.ContNo=*ImportCont.ContNo AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl AND " & _
          "BL.BLNo=*ImportCont.BLNo AND ImportCargoInfo.ContNo=*ImportCont.ContNo AND ImportCargoInfo.VslCode=*ImportCont.VslCode AND ImportCargoInfo.BLNo=*ImportCont.BLNo AND " & _
          "VslSchedule.VslCode=*ImportCont.VslCode AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl AND " & _
          "mShipLine.ShipLine=*Container.ShipLine AND mVessel.VslCode=*ImportCont.VslCode AND " & _
          "VslPort.VslCode=*ImportCont.VslCode AND VslPort.VoyCtrl=*ImportCont.VoyCtrl AND VslPort.PortCode='" & strPort & "'"
    sql = sql & " ORDER BY ImportCont.BLNo, ImportCont.ContNo"
    'SQLを発行して検索
    rsd.Open sql, conn, 0, 1, 1

    Dim sOutText()

    Do While Not rsd.EOF
        '日付までしか入っていないDateTime型の比較準備(2003/10/25 Add)
        strchkNow = DispDateTime( Now, 8 )
        strchkFreeTime = DispDateTime( rsd("FreeTime"), 8 )
        strchkDelPermitDate = DispDateTime( rsd("DelPermitDate"), 8 )
        strchkOLTDateFrom = DispDateTime( rsd("OLTDateFrom"), 8 )
        strchkOLTDateTo = DispDateTime( rsd("OLTDateTo"), 8 )
        strchkFreeTimeExt = DispDateTime( rsd("FreeTimeExt"), 8 )
        '日付までしか入っていないDateTime型の比較準備(2003/10/25)
        ' 通関／保税輸送のチェック
        If Trim(rsd("OLTICFlag"))="I" Then
            If Trim(rsd("OLTICNo"))<>"" Then
                strTsukan = "I"
            Else
                strTsukan = ""
            End If
        Else
            '日付までしか入っていないDateTime型の比較に変更(2003/10/25 Chenge)
'            If Now>=rsd("OLTDateFrom") And Now<=rsd("OLTDateTo") Then
            If strchkNow>=strchkOLTDateFrom And strchkNow<=strchkOLTDateTo Then
           '日付までしか入っていないDateTime型の比較に変更(2003/10/25)
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

	'船名のチェックを止める(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ",,"
'        Else
            strOut = strOut & DispDateTime(rsd("ETA"),0) & "," ' 3:着岸予定時刻
            strOut = strOut & DispDateTime(rsd("TA"),0) & ","  ' 4:着岸完了時刻
'        End If
	'船名のチェックを止める(2004/03/12)
        ' 5:搬出可否
        If DispDateTime(rsd("CYDelTime"),0)<>"" Then
            strOut = strOut & "S,"
        '日付までしか入っていないDateTime型の比較に変更(2003/10/25 Chenge)
'        ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
'           (rsd("FreeTime")>=Now Or rsd("FreeTimeExt")>=Now) And _
'           rsd("DelPermitDate")<=Now Then
        ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
           (strchkFreeTime>=strchkNow Or strchkFreeTimeExt>=strchkNow) And _
           strchkDelPermitDate<=strchkNow Then
            '日付までしか入っていないDateTime型の比較に変更(2003/10/25)
            strOut = strOut & "Y,"
        Else
            strOut = strOut & ","
        End If
'        strOut = strOut & Trim(rsd("StackPlace")) & ","        ' 6:搬出ターミナル名
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
'            strOut = strOut & DispDateTime(rsd("TD"),0) & ","     '12:離岸完了時刻
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
        '28:危険品
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
	'船名のチェックを止める(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ","
'        Else
            strOut = strOut & DispDateTime(rsd("ETALong"),0) & ","       '32:着岸スケジュール
'        End If
	'船名のチェックを止める(2004/03/12)
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
        If strWHA="" Then
            strWHA = DispDateTime(rsd("WHArTime"),0)
        End If
        strOut = strOut & strWHA & ","
        strOut = strOut & Trim(rsd("SYDelFlag")) & ","                   '36:ストックヤード搬出予約フラグ
        strWork=DispDateTime(rsd("SYPRDelTime"),0)
        If Len(strWork)>13 Then
            strTemp=Right("0" & CInt(Mid(strWork, 12, 2))+1, 2)
            strWork=Left(strWork, 11) & strTemp & Right(strWork, 3)
        End If
        strOut = strOut & strWork & ","                                  '37:ストックヤード搬出予約時刻

        ReDim Preserve sOutText(bWrite)
        sOutText(bWrite) = strOut
        bWrite = bWrite + 1

        rsd.MoveNext
    Loop
    rsd.Close

    '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2003/09/24 Add)
    iNotWriteRecords = 0
    '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2003/09/24)
    For i=0 To bWrite-1
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2003/09/24 Add)
        If sOutText(i)<>"" Then
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2003/09/24)
        strTmp=Split(sOutText(i),",")
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2003/09/24 Add)
        '同一コンテナが複数あるとき、一つにまとめる
        For j=i+1 To bWrite-1
            If sOutText(j)<>"" Then             '(2004/07/09)
            strTmp2=Split(sOutText(j),",")
            If strTmp(1)=strTmp2(1) Then
                sOutText(j)=""
            End If
            End If                              '(2004/07/09)
        Next

        'コンテナ単位に読み直してチェックする。１BLでも通関付加があれば、搬出を不許可とする
        sOutFlag = "Y"
        sql = "SELECT OLTICFlag, OLTICNo, OLTDateFrom, OLTDateTo, CYDelTime, DOStatus, FreeTime, FreeTimeExt, DelPermitDate FROM ImportCont " & _
              "WHERE ContNo='" & strTmp(1) & "'"
        'SQLを発行して検索
        rsd.Open sql, conn, 0, 1, 1

        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2004/02/17 Add)
        'If Not rsd.EOF Then
        Do While Not rsd.EOF
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2004/02/17)
            '日付までしか入っていないDateTime型の比較準備(2003/10/25 Add)
            strchkNow = DispDateTime( Now, 8 )
            strchkFreeTime = DispDateTime( rsd("FreeTime"), 8 )
            strchkDelPermitDate = DispDateTime( rsd("DelPermitDate"), 8 )
            strchkOLTDateFrom = DispDateTime( rsd("OLTDateFrom"), 8 )
            strchkOLTDateTo = DispDateTime( rsd("OLTDateTo"), 8 )
            strchkFreeTimeExt = DispDateTime( rsd("FreeTimeExt"), 8 )
            '日付までしか入っていないDateTime型の比較準備(2003/10/25)
            ' 通関／保税輸送のチェック
            If Trim(rsd("OLTICFlag"))="I" Then
                If Trim(rsd("OLTICNo"))<>"" Then
                    strTsukan = "I"
                Else
                    strTsukan = ""
                End If
            Else
                '日付までしか入っていないDateTime型の比較に変更(2003/10/25 Chenge)
'                If Now>=rsd("OLTDateFrom") And Now<=rsd("OLTDateTo") Then
                If strchkNow>=strchkOLTDateFrom And strchkNow<=strchkOLTDateTo Then
                '日付までしか入っていないDateTime型の比較に変更(2003/10/25)
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
            ' 5:搬出可否
            If DispDateTime(rsd("CYDelTime"),0)<>"" Then
                sOutFlag = "S"                                      '搬出済
            '日付までしか入っていないDateTime型の比較に変更(2003/10/25 Chenge)
'            ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
'               (rsd("FreeTime")>=Now Or rsd("FreeTimeExt")>=Now) And _
'               rsd("DelPermitDate")<=Now Then
            ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
               (strchkFreeTime>=strchkNow Or strchkFreeTimeExt>=strchkNow) And _
               strchkDelPermitDate<=strchkNow Then
                '日付までしか入っていないDateTime型の比較に変更(2003/10/25)
                'フラグ変更無し
            Else
                sOutFlag = ""                                       '搬出不可
                '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2004/02/17 Add)
                Exit Do
                '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2004/02/17)
            End If
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2004/02/17 Add)
        'End If
            rsd.MoveNext
        Loop
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2004/02/17)
        rsd.Close

        If sOutFlag<>strTmp(4) Then
            'ﾃﾞｰﾀ変更反映
            strTmp(4) = sOutFlag
            strOut = ""
            For j=0 To 36
                strOut = strOut + strTmp(j) + ","
            Next
            sOutText(i) = strOut
        End If
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2003/09/24)

        ' 積港データを探す
        If strTmp(9)<>"" Then
            sql = "SELECT ImportCont.VoyCtrl, VslPort.TD FROM ImportCont, VslPort " & _
                  "WHERE ImportCont.ContNo='" & strTmp(1) & "' AND ImportCont.BLNo='" & strTmp(0) & "' AND VslPort.VslCode=ImportCont.VslCode AND " & _
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
                strWork = Trim(rsd("FullName"))
                If ( Not IsNull(strWork) ) And strWork<>"" Then
                    strTmp(9) = Trim(rsd("FullName"))          '10:仕出港
                End If
            End If
            rsd.Close

            strWork=""
            For j=0 To UBound(strTmp)-1
                strWork=strWork & strTmp(j) & ","
            Next
            sOutText(i)=strWork
        End If

        ' 前港データを探す
        sql = "SELECT ImportCont.VoyCtrl, B.TD, mPort.FullName FROM ImportCont, VslPort A, VslPort B, mPort " & _
              "WHERE ImportCont.ContNo='" & strTmp(1) & "' AND ImportCont.BLNo='" & strTmp(0) & "' AND A.VslCode=ImportCont.VslCode AND " & _
              "A.VoyCtrl=ImportCont.VoyCtrl AND A.PortCode='" & strPort & "' AND " & _
              "B.VslCode=ImportCont.VslCode AND B.VoyCtrl=ImportCont.VoyCtrl AND B.CallSeq=(A.CallSeq-1) AND " & _
              "mPort.PortCode=*B.PortCode"
        'SQLを発行して検索
        rsd.Open sql, conn, 0, 1, 1

        If Not rsd.EOF Then
            If strTmp(7)<>"" Then    '船名があったら
                sOutText(i) = sOutText(i) & DispDateTime(rsd("TD"),0) & ","    '38:前港離岸完了
            Else
                sOutText(i) = sOutText(i) & ","
            End If
            sOutText(i) = sOutText(i) & Trim(rsd("FullName"))                  '39:前港
        Else
            sOutText(i) = sOutText(i) & ","
        End If
        rsd.Close
'HiTS ver2 ADD START by SEIKO n.Ooshige 2003/06/26
	'ディテンションフリータイムと事前入力作業番号の取得
	sql="SELECT IPC.DetentionFreeTime, ITC1.WkNo AS WkNo1, ITC2.WkNo AS WkNo2, ITC3.WkNo AS WkNo3 "&_
            "FROM (((ImportCont AS IPC LEFT JOIN hITCommonInfo AS ITC1 ON IPC.ContNo = ITC1.ContNo) "&_
            "LEFT JOIN hITFullOutSelect AS ITF ON IPC.ContNo = ITF.ContNo) "&_
            "LEFT JOIN hITCommonInfo AS ITC2 ON ITF.WkContrlNo = ITC2.WkContrlNo) "&_
            "LEFT JOIN hITCommonInfo AS ITC3 ON IPC.BLNo = ITC3.BLNo "&_
            "WHERE IPC.ContNo='"& strTmp(1) &"'"
        rsd.Open sql, conn, 0, 1, 1
        sOutText(i) = sOutText(i)&","& Trim(rsd("DetentionFreeTime"))		'40:ディテンションフリータイム
        If Not IsNull(rsd("WkNo1")) Then
            sOutText(i) = sOutText(i)&","& Trim(rsd("WkNo1"))			'41:事前入力作業番号
        ElseIf Not IsNull(rsd("WkNo2")) Then
            sOutText(i) = sOutText(i)&","& Trim(rsd("WkNo2"))			'41:事前入力作業番号
        Else
            sOutText(i) = sOutText(i)&","& Trim(rsd("WkNo3"))			'41:事前入力作業番号
        End If
        rsd.Close
'HiTS ver2 ADD END   by SEIKO n.Ooshige 2003/06/26

        pfp.WriteLine sOutText(i)
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2003/09/24 Add)
        Else
            iNotWriteRecords = iNotWriteRecords + 1
        End If
        '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2003/09/24)
    Next

    '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮する(2003/09/24 Add)
''    SerchImpCntnr = bWrite    '該当レコードあったかどうか
    SerchImpCntnr = bWrite - iNotWriteRecords  '該当レコードあったかどうか
    '輸入コンテナ搬出可否のチェックにおいて、複数BLを考慮するここまで(2003/09/24)
End Function

'輸入コンテナCSVファイルタイトル行出力
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]検索モード(true=コンテナ検索 / false=BL検索)

    Response.Write "BL番号,"
    Response.Write "コンテナNo.,"

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
    Response.Write "危険品,"
    Response.Write "搬出ターミナル名,"
    Response.Write "ストックヤード利用,"
    Response.Write "返却先,"

    Response.Write "船社,"
    Response.Write "船名,"
    Response.Write "Voyage No.,"
    Response.Write "仕出港,"
    Response.Write "前港"

'    Response.Write "保税輸送(From),"
'    Response.Write "保税輸送(To)"

'HiTS ver2 ADD START by SEIKO n.Ooshige 2003/07/09
    Response.Write ",ディテンションフリータイム"
    Response.Write ",事前入力作業番号"
'HiTS ver2 ADD END by SEIKO n.Ooshige 2003/07/09

    Response.Write Chr(13) & Chr(10)
End Sub

'輸入コンテナCSVファイルデータ行出力
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]検索モード(true=コンテナ検索 / false=BL検索)
    'pfp        [ I ]テンポラリファイル

    ' ダウンロードファイルのレコードがある間繰り返す
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        ' BL番号
        If Not bMode Then
            Response.Write anyTmp(0) & ","
        Else
            Response.Write ","
        End If
        ' コンテナNo.
        Response.Write anyTmp(1) & ","

        ' 離岸完了時刻(仕出港)
        Response.Write anyTmp(11) & ","
        ' 離岸完了時刻(前港)
        Response.Write anyTmp(37) & ","
        ' 着岸スケジュール
        Response.Write Left(anyTmp(31),10) & ","
        ' 着岸予定時刻
        Response.Write anyTmp(2) & ","
        ' 着岸完了時刻
        Response.Write anyTmp(3) & ","
        ' ヤード搬入
        Response.Write anyTmp(12) & ","
        ' ヤード搬出
        Response.Write anyTmp(13) & ","
        ' ストックヤード搬出予約
        sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
        If anyTmp(35)>="1" And anyTmp(35)<="4" Then
            sTemp=Left(anyTmp(36),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' ストックヤード搬出
        Response.Write anyTmp(30) & ","
        ' 倉庫到着スケジュール
        Response.Write anyTmp(34) & ","
        ' 倉庫到着完了時刻
        Response.Write anyTmp(14) & ","
        ' デバン
        Response.Write anyTmp(15) & ","
        ' 空コン返却
        Response.Write anyTmp(16) & ","

        ' 搬入確認予定時刻
        Response.Write anyTmp(32) & ","
        ' 搬入確認完了時刻
        Response.Write Left(anyTmp(18),10) & ","
        ' 動植物
        If anyTmp(17)="S" Then
            Response.Write "×,"
        ElseIf anyTmp(17)="C" Then
            Response.Write "○,"
        Else
            Response.Write "−,"
        End If
        ' 個別搬入
        If anyTmp(33)<>"" Then
            Response.Write "○,"
        Else
            Response.Write "×,"
        End If
        ' 通関／保税輸送
        If anyTmp(19)<>"" Then
            Response.Write "○,"
        Else
            Response.Write "×,"
        End If
        ' 商取引DO発行
        If anyTmp(21)<>"Y" Then
            Response.Write "×,"
        Else
            Response.Write "○,"
        End If
        ' フリータイム
        Response.Write Left(anyTmp(22),10) & ","
        ' 搬出可否
        If anyTmp(4)="Y" Then
            Response.Write "○,"
        ElseIf anyTmp(4)="S" Then
            Response.Write "済,"
        Else
            Response.Write "×,"
        End If

        ' サイズ
        Response.Write anyTmp(23) & ","
        ' 高さ
        Response.Write anyTmp(24) & ","
        ' リーファー
        If anyTmp(25)="R" Then
            Response.Write "○,"
        ElseIf anyTmp(25)<>"" Then
            Response.Write "−,"
        Else
            Response.Write ","
        End If
        ' 重量
        If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
            dWeight=anyTmp(26) / 10
            Response.Write dWeight
        Else
            Response.Write "−"
        End If
        Response.Write ","
        ' 危険品
        If anyTmp(27)="H" Then
            Response.Write "○,"
        ElseIf anyTmp(27)<>"" Then
            Response.Write "−,"
        Else
            Response.Write ","
        End If
        ' 搬出ターミナル名
        Response.Write anyTmp(5) & ","
        ' ストックヤード利用
        If anyTmp(35)>="1" And anyTmp(35)<="4" Then
            Response.Write "○,"
        Else
            Response.Write "×,"
        End If
        ' 返却場所
        Response.Write anyTmp(10) & ","

        ' 船社名
        Response.Write anyTmp(6) & ","
        ' 船名
        Response.Write anyTmp(7) & ","
        ' 次航
        Response.Write anyTmp(8) & ","
        ' 仕出港
        Response.Write anyTmp(9) & ","
        ' 前港
        Response.Write anyTmp(38)

        ' 保税輸送(From)
'        Response.Write anyTmp(28) & ","
        ' 保税輸送(To)
'        Response.Write anyTmp(29)

'HiTS ver2 ADD START by SEIKO n.Ooshige 2003/07/09
	'ディテンションフリータイム
        Response.Write "," & anyTmp(39)
	'事前入力作業番号
        Response.Write "," & anyTmp(40)
'HiTS ver2 ADD END by SEIKO n.Ooshige 2003/07/09

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>