<%
'�A���Ɖ�֌W�֐�

'�A���R���e�i�Ɖ�
Function SerchMSImpCntnr(conn, rsd, pfp, sWhere, sOrder)
    'conn       [ I ]DB�ڑ�Object
    'rsd        [ I ]���R�[�h�Z�b�gObject
    'pfp        [ I ]�e���|�����t�@�C��
    'sWhere     [ I ]��������
    'sOrder     [ I ]�\�[�g����

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    bWrite = 0        '�o�̓��R�[�h����

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

    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1

    Do While Not rsd.EOF
        sText=""

        ' �o�̓��R�[�h�̕ҏW
        sText = Trim(rsd("BLNo")) & ","                      ' 1:BL�ԍ�
        sText = sText & Trim(rsd("ContNo")) & ","            ' 2:�R���e�iNo.
        sWork = Trim(rsd("ShipFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("VslCode")) & ","       ' 3:�D��(Full)
        Else
            sText = sText & sWork & ","                      ' 3:�D��
        End If
        sText = sText & Trim(rsd("DsVoyage")) & ","          ' 4:���q
        sText = sText & rsd("VslCode") & ","                 ' 5:�D��
        sText = sText & rsd("VoyCtrl") & ","                 ' 6:���q(�ԍ�)
        sWork = Trim(rsd("ShipName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("VslCode")) & ","       ' 7:�D��(����)
        Else
            sText = sText & sWork & ","                      ' 7:�D��(����)
        End If
        sWork = Trim(rsd("ShipperName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Shipper")) & ","       ' 8:�׎�(����)
        Else
            sText = sText & sWork & ","                      ' 8:�׎�(����)
        End If
        sWork = Trim(rsd("ForName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Forwarder")) & ","     ' 9:�C��(����)
        Else
            sText = sText & sWork & ","                      ' 9:�C��(����)
        End If
        sWork = Trim(rsd("TruckName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Trucker")) & ","       '10:���^(����)
        Else
            sText = sText & sWork & ","                      '10:���^(����)
        End If
        sText = sText & Trim(rsd("ContSize")) & ","          '11:�T�C�Y
        sText = sText & Trim(rsd("ContType")) & ","          '12:�^�C�v
        sText = sText & Trim(rsd("Remark")) & ","            '13:�R�����g
        sText = sText & DispDateTime(rsd("WHArTime"),0) & ","     '14:�q�ɓ����w�����
        sText = sText & Trim(rsd("TruckRefFlag")) & ","      '15:���^�Q�ƃt���O
        sText = sText & Trim(rsd("ShipLine")) & ","          '16:�D�ЃR�[�h
        sText = sText & Trim(rsd("BLVoyage")) & ","          '17:���q(�ԍ�)
        sWork = Trim(rsd("ShipperFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Shipper")) & ","       '18:�׎�(�t������)
        Else
            sText = sText & sWork & ","                      '18:�׎�(�t������)
        End If
        sWork = Trim(rsd("ForFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Forwarder")) & ","     '19:�C��(�t������)
        Else
            sText = sText & sWork & ","                      '19:�C��(�t������)
        End If
        sWork = Trim(rsd("TruckFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("Trucker")) & ","       '20:���^(�t������)
        Else
            sText = sText & sWork & ","                      '20:���^(�t������)
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
            'SQL�𔭍s���Č���
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

            'SQL�𔭍s���Č���
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' �ʊց^�ېŗA���̃`�F�b�N
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
                If DispDateTime(rsd("CYDelTime"),0)<>"" Then           ' ���o����Ă����灛�Ƃ���
                    If IsNull(rsd("OLTDateFrom")) Or IsNull(rsd("OLTDateTo")) Then
                        strTsukan = "S"
                    Else
                        strTsukan = "T"
                    End If
                End If

                ' �o�̓��R�[�h�̕ҏW
                strOut = Trim(rsd("BLNo")) & ","                       ' 1:BL�ԍ�
                strOut = strOut & Trim(rsd("ContNo")) & ","            ' 2:�R���e�iNo.

                strCheck=rsd("ShipName")                '�D�����Ȃ�������A���݁E���ݎ�����\�����Ȃ��悤�ɂ���

                If IsNull( strCheck ) Then
                    strOut = strOut & ",,"
                Else
                    strOut = strOut & DispDateTime(rsd("ETA"),0) & "," ' 3:���ݗ\�莞��
                    strOut = strOut & DispDateTime(rsd("TA"),0) & ","  ' 4:���݊�������
                End If
                ' 5:���o��
                If DispDateTime(rsd("CYDelTime"),0)<>"" Then
                    strOut = strOut & "S,"
                ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
                   (rsd("FreeTime")>=Now Or rsd("FreeTimeExt")>=Now) And _
                   rsd("DelPermitDate")<=Now Then
                    strOut = strOut & "Y,"
                Else
                    strOut = strOut & ","
                End If
'                strOut = strOut & Trim(rsd("StackPlace")) & ","        ' 6:���o�^�[�~�i����
                strOut = strOut & Trim(rsd("RecTerminal")) & ","       ' 6:���o�^�[�~�i����

                If IsNull(rsd("ShipLineName")) Then
                    strOut = strOut & Trim(rsd("ShipLine")) & ","      ' 7:�D�Ж�
                Else
                    strOut = strOut & Trim(rsd("ShipLineName")) & ","  ' 7:�D�Ж�
                End If
                If IsNull(rsd("ShipName")) Then
                    strOut = strOut & Trim(rsd("VslCode")) & ","       ' 8:�D��
                Else
                    strOut = strOut & Trim(rsd("ShipName")) & ","      ' 8:�D��
                End If
                strOut = strOut & Trim(rsd("DsVoyage")) & ","          ' 9:���q
                strOut = strOut & Trim(rsd("LPort")) & ","             '10:�d�o�`(�R�[�h)
                strOut = strOut & Trim(rsd("ReturnPlace")) & ","       '11:�ԋp�ꏊ

                ' ���n��̃`�F�b�N
                strCYDelTime = DispDateTime(rsd("CYDelTime"),0)        '14:���[�h���o
                strSYDelTime = DispDateTime(rsd("SYDelTime"),0)        '31:�X�g�b�N���[�h���o����
                strWHArDate = DispDateTime(rsd("WHArDate"),0)          '15:�q�ɓ���
                strDevanTime = DispDateTime(rsd("DevanTime"),0)        '16:�f�o��
                strReturnTime = DispDateTime(rsd("ReturnTime"),0)      '17:��R���ԋp
                strOld = strSYDelTime
                If strOld = "" Then
                    strOld = strCYDelTime
                End If
                If strWHArDate<strOld Then
                    strWHArDate = ""                      ' �q�ɓ�����KACCS,Hits�������ߋ��̂Ƃ��\�����Ȃ�
                End If
                If strReturnTime<>"" And strWHArDate>strReturnTime Then
                    strWHArDate = ""                      ' �q�ɓ�������R���ԋp������薢���̂Ƃ��\�����Ȃ�
                End If
                If strWHArDate<>"" Then
                    strOld = strWHArDate
                End If
                If strDevanTime<strOld Then
                    strDevanTime = ""                     ' �f�o����KACCS,Hits�������ߋ��̂Ƃ��\�����Ȃ�
                End If
                If strReturnTime<>"" And strDevanTime>strReturnTime Then
                    strDevanTime = ""                     ' �f�o������R���ԋp������薢���̂Ƃ��\�����Ȃ�
                End If

                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & ","
'                    strOut = strOut & DispDateTime(rsd("TD"),0) & ","     '12:���݊�������
                End If
                strOut = strOut & DispDateTime(rsd("InTime"),0) & ","     '13:���[�h����
                strOut = strOut & DispDateTime(rsd("CYDelTime"),0) & ","  '14:���[�h���o
                strOut = strOut & strWHArDate & ","                       '15:�q�ɓ���
                strOut = strOut & strDevanTime & ","                      '16:�f�o��
                strOut = strOut & DispDateTime(rsd("ReturnTime"),0) & "," '17:��R���ԋp

                '18:���A��
                strBotStat=Trim(rsd("BotStatus"))
                strZooStat=Trim(rsd("ZooStatus"))
                If strBotStat="S" Or strBotStat="C" Then
                    strOut = strOut & strBotStat & ","
                ElseIf strZooStat="S" Or strZooStat="C" Then
                    strOut = strOut & strZooStat & ","
                Else
                    strOut = strOut & ","
                End If
                strOut = strOut & DispDateTime(rsd("ListDate"),0) & ","    '19:�����m�F��������
                strOut = strOut & strTsukan & ","                          '20:�ʊց^�ېŗA��
                strOut = strOut & Trim(rsd("CustStatus")) & ","            '21:�ʊ֐R����
                strOut = strOut & Trim(rsd("DOStatus")) & ","              '22:�����DO���s
                '23:�t���[�^�C��
                strFree = DispDateTime(rsd("FreeTime"),10)
                strFreeExt = DispDateTime(rsd("FreeTimeExt"),10)
                If strFreeExt<>"" Then
                    strFree = strFreeExt
                End If
                strOut = strOut & strFree & ","

                strOut = strOut & Trim(rsd("ContSize")) & ","          '24:�T�C�Y
                strOut = strOut & Trim(rsd("ContHeight")) & ","        '25:����
                '26:���[�t�@�[
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
                strOut = strOut & Trim(rsd("ContWeight")) & ","        '27:�d��
                '28:�댯��
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

                strOut = strOut & DispDateTime(rsd("OLTDateFrom"),10) & "," '29:�ېŗA��(From)
                strOut = strOut & DispDateTime(rsd("OLTDateTo"),10) & ","   '30:�ېŗA��(To)

                strOut = strOut & DispDateTime(rsd("SYDelTime"),0) & ","         '31:�X�g�b�N���[�h���o����
                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & DispDateTime(rsd("ETALong"),0) & ","       '32:���݃X�P�W���[��
                End If
                '33:�����m�F�\�莞��
                strIndv = DispDateTime(rsd("IndvRecSchTime"),0)
                If strIndv="" Then
                    strIndv = DispDateTime(rsd("BatchRecSchTime"),0)
                End If
                strOut = strOut & strIndv & ","
                strOut = strOut & Trim(rsd("IndvRecFlag")) & ","                 '34:�ʔ���
                '35:�q�ɓ����X�P�W���[��
                strWHA = DispDateTime(rsd("ImCntnrWHArSch"),0)
                If strWHA="" Then
                    strWHA = DispDateTime(rsd("BLWHArSch"),0)
                End If
                strOut = strOut & strWHA & ","
                strOut = strOut & Trim(rsd("SYDelFlag")) & ","                   '36:�X�g�b�N���[�h���o�\��t���O
                strWork=DispDateTime(rsd("SYPRDelTime"),0)
                If Len(strWork)>13 Then
                    strTemp=Right("0" & CInt(Mid(strWork, 12, 2))+1, 2)
                    strWork=Left(strWork, 11) & strTemp & Right(strWork, 3)
                End If
                strOut = strOut & strWork & ","                                  '37:�X�g�b�N���[�h���o�\�񎞍�

                rsd.Close

                ' �ύ`�f�[�^��T��
                strTmp=Split(strOut,",")
                If strTmp(9)<>"" Then
                    sql = "SELECT ImportCont.VoyCtrl, VslPort.TD FROM ImportCont, VslPort " & _
                          "WHERE (" & sWhere & ") AND VslPort.VslCode=ImportCont.VslCode AND " & _
                          "VslPort.VoyCtrl=ImportCont.VoyCtrl AND VslPort.PortCode='" & strTmp(9) & "'"
                    'SQL�𔭍s���Č���
                    rsd.Open sql, conn, 0, 1, 1

                    If Not rsd.EOF Then
                        If strTmp(7)<>"" Then    '�D������������
                            strTmp(11) = DispDateTime(rsd("TD"),0)     '12:���݊�������
                        End If
                    End If
                    rsd.Close

                    sql = "SELECT mPort.FullName From mPort " & _
                          "WHERE mPort.PortCode='" & strTmp(9) & "'"
                    'SQL�𔭍s���Č���
                    rsd.Open sql, conn, 0, 1, 1

                    If Not rsd.EOF Then
                        If Not IsNull(rsd("FullName")) Then
                            strTmp(9) = Trim(rsd("FullName"))          '10:�d�o�`
                        End If
                    End If
                    rsd.Close

                    strWork=""
                    For j=0 To UBound(strTmp)-1
                        strWork=strWork & strTmp(j) & ","
                    Next
                    strOut=strWork
                End If

                ' �O�`�f�[�^��T��
                sql = "SELECT ImportCont.VoyCtrl, B.TD, mPort.FullName FROM ImportCont, VslPort A, VslPort B, mPort " & _
                      "WHERE (" & sWhere & ") AND A.VslCode=ImportCont.VslCode AND " & _
                      "A.VoyCtrl=ImportCont.VoyCtrl AND A.PortCode='" & strPort & "' AND " & _
                      "B.VslCode=ImportCont.VslCode AND B.VoyCtrl=ImportCont.VoyCtrl AND B.CallSeq=(A.CallSeq-1) AND " & _
                      "mPort.PortCode=*B.PortCode"
                'SQL�𔭍s���Č���
                rsd.Open sql, conn, 0, 1, 1

                If Not rsd.EOF Then
                    If IsNull( strCheck ) Then
                        strOut = strOut & ","
                    Else
                        strOut = strOut & DispDateTime(rsd("TD"),0) & ","      '38:�O�`���݊���
                    End If
                    strOut = strOut & Trim(rsd("FullName"))                    '39:�O�`
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

            'SQL�𔭍s���Č���
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' �o�̓��R�[�h�̕ҏW
                strOut = Trim(rsd("BLNo")) & ","                       ' 1:BL�ԍ�
                strOut = strOut & strTmp(1) & ","                      ' 2:�R���e�iNo.
                strOut = strOut & DispDateTime(rsd("ETA"),0) & ","     ' 3:���ݗ\�莞��
                strOut = strOut & DispDateTime(rsd("TA"),0) & ","      ' 4:���݊�������
                strOut = strOut & ","                                  ' 5:���o��
                strOut = strOut &  Trim(rsd("RecTerminal")) & ","      ' 6:���o�^�[�~�i����

                strOut = strOut & Trim(rsd("ShipLineName")) & ","      ' 7:�D�Ж�
                strOut = strOut & Trim(rsd("ShipName")) & ","          ' 8:�D��
                strOut = strOut & Trim(rsd("DsVoyage")) & ","          ' 9:���q
                strOut = strOut & Trim(rsd("LPortName")) & ","         '10:�d�o�`
                strOut = strOut & ","                                  '11:�ԋp�ꏊ

                strOut = strOut & DispDateTime(rsd("TD"),0) & ","      '12:���݊�������
                strOut = strOut & ","                                  '13:���[�h����
                strOut = strOut & ","                                  '14:���[�h���o
                strOut = strOut & ","                                  '15:�q�ɓ���
                strOut = strOut & ","                                  '16:�f�o��
                strOut = strOut & ","                                  '17:��R���ԋp

                strOut = strOut & ","                                  '18:���A��
                strOut = strOut & ","                                  '19:�����m�F��������
                strOut = strOut & ","                                  '20:�ʊց^�ېŗA��
                strOut = strOut & ","                                  '21:�ʊ֐R����
                strOut = strOut & ","                                  '22:�����DO���s
                strOut = strOut & ","                                  '23:�t���[�^�C��

                strOut = strOut & ","                                  '24:�T�C�Y
                strOut = strOut & ","                                  '25:����
                strOut = strOut & Trim(rsd("RFFlag")) & ","            '26:���[�t�@�[
                strOut = strOut & ","                                  '27:�d��
                strOut = strOut & Trim(rsd("DGFlag")) & ","            '28:�댯��

                strOut = strOut & ","                                  '29:�ېŗA��(From)
                strOut = strOut & ","                                  '30:�ېŗA��(To)

                strOut = strOut & ","                                  '31:�X�g�b�N���[�h���o����
                strOut = strOut & DispDateTime(rsd("ETALong"),0) & ","           '32:���݃X�P�W���[��
                strOut = strOut & DispDateTime(rsd("IndvRecSchTime"),0) & ","    '33:�����m�F�\�莞��
                strOut = strOut & Trim(rsd("IndvRecFlag")) & ","                 '34:�ʔ���
                strOut = strOut & DispDateTime(rsd("BLWHArSch"),0) & ","         '35:�q�ɓ����X�P�W���[��
                strOut = strOut & ","                                  '36:�X�g�b�N���[�h���o�\��t���O
                strOut = strOut & ","                                  '37:�X�g�b�N���[�h���o�\�񎞍�
                strOut = strOut & DispDateTime(rsd("BeforeTD"),0) & ","          '38:�O�`���݊���
                strOut = strOut & Trim(rsd("BeforePort"))                        '39:�O�`
            Else
                strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            rsd.Close
        Else
            strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
        End If

        pfp.WriteLine Tmp(i) & strOut & ","
    Next

    SerchMSImpCntnr = bWrite    '�Y�����R�[�h���������ǂ���
End Function

'�A���R���e�iCSV�t�@�C���^�C�g���s�o��
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]�������[�h(���O�C�����)

    If bMode="�C��" Then
        Response.Write "�D��,"
        Response.Write "Voyage No.,"
        Response.Write "�׎�,"
        Response.Write "�D��,"
    ElseIf bMode="���^" Then
        Response.Write "�C��,"
    Else
        Response.Write "�D��,"
        Response.Write "Voyage No.,"
        Response.Write "�C��,"
        Response.Write "�D��,"
    End If

    Response.Write "BL�ԍ�,"
    Response.Write "�R���e�iNo.,"

    If bMode="�C��" Then
        Response.Write "�w�藤�^�Ǝ�,"
    ElseIf bMode="���^" Then
        Response.Write "�q�ɗ���,"
    Else
        Response.Write "�w�藤�^�Ǝ�,"
    End If

    Response.Write "�d�o�`���݊�������,"
    Response.Write "�O�`���݊�������,"
    Response.Write "CY���݌v��,"
    Response.Write "CY���ݗ\�莞��,"
    Response.Write "CY���݊�������,"
    Response.Write "CY������������,"
    Response.Write "CY���o��������,"
    Response.Write "SY�\�񎞍�,"
    Response.Write "SY���o��������,"
    Response.Write "�q�ɓ����w������,"
    Response.Write "�q�ɓ�����������,"
    Response.Write "�f�o����������,"
    Response.Write "��R���ԋp����,"

    Response.Write "�����m�F�\�莞��,"
    Response.Write "�����m�F��������,"
    Response.Write "���A�����u,"
    Response.Write "�ʔ���,"
    Response.Write "�ʊ�/�ېŗA��,"
    Response.Write "DO���s,"
    Response.Write "�t���[�^�C��,"
    Response.Write "���o��,"

    Response.Write "�T�C�Y,"
    Response.Write "����,"
    Response.Write "���[�t�@�[,"
    Response.Write "���d��,"
    Response.Write "�댯��,"
    Response.Write "���o�^�[�~�i����,"
    Response.Write "�X�g�b�N���[�h���p,"
    Response.Write "�ԋp��,"

    If bMode="���^" Then
        Response.Write "�D��,"
        Response.Write "�D��,"
        Response.Write "Voyage No.,"
    End If
    Response.Write "�d�o�`,"
    Response.Write "�O�`"

'    Response.Write "�ېŗA��(From),"
'    Response.Write "�ېŗA��(To)"

    Response.Write Chr(13) & Chr(10)
End Sub

'�A���R���e�iCSV�t�@�C���f�[�^�s�o��
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]�������[�h(���O�C�����)
    'pfp        [ I ]�e���|�����t�@�C��

    ' �_�E�����[�h�t�@�C���̃��R�[�h������ԌJ��Ԃ�
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        If bMode="�C��" Then
            ' �D��
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
            ' �׎�
            Response.Write anyTmp(17) & ","
            ' �D��
            If anyTmp(36)<>"" Then
                Response.Write anyTmp(36) & ","
            ElseIf anyTmp(21)<>"" Then
                Response.Write anyTmp(21) & ","
            Else
                Response.Write anyTmp(15) & ","
            End If
        ElseIf bMode="���^" Then
            ' �C��
            Response.Write anyTmp(18) & ","
        Else
            ' �D��
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
            ' �C��
            Response.Write anyTmp(18) & ","
            ' �D��
            If anyTmp(36)<>"" Then
                Response.Write anyTmp(36) & ","
            ElseIf anyTmp(21)<>"" Then
                Response.Write anyTmp(21) & ","
            Else
                Response.Write anyTmp(15) & ","
            End If
        End If

        ' BL�ԍ�
        Response.Write anyTmp(0) & ","
        ' �R���e�iNo.
        Response.Write anyTmp(1) & ","

        If bMode="�C��" Then
            ' �w�藤�^�Ǝ�
            Response.Write anyTmp(19) & ","
        ElseIf bMode="���^" Then
            ' �q�ɗ���
            Response.Write anyTmp(12) & ","
        Else
            ' �w�藤�^�Ǝ�
            Response.Write anyTmp(19) & ","
        End If

        ' ���݊�������(�d�o�`)
        Response.Write anyTmp(41) & ","
        ' ���݊�������(�O�`)
        Response.Write anyTmp(67) & ","
        ' ���݃X�P�W���[��
        Response.Write Left(anyTmp(61),10) & ","
        ' ���ݗ\�莞��
        Response.Write anyTmp(32) & ","
        ' ���݊�������
        Response.Write anyTmp(33) & ","
        ' ���[�h����
        Response.Write anyTmp(42) & ","
        ' ���[�h���o
        Response.Write anyTmp(43) & ","
        ' �X�g�b�N���[�h���o�\��
        sTemp=DispReserveCell(anyTmp(65),anyTmp(66),sColor)
        If anyTmp(65)>="1" And anyTmp(65)<="4" Then
            sTemp=Left(anyTmp(66),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' �X�g�b�N���[�h���o
        Response.Write anyTmp(60) & ","
        ' �q�ɓ����X�P�W���[��
        If anyTmp(64)<>"" Then
            Response.Write anyTmp(64) & ","
        Else
            Response.Write anyTmp(13) & ","
        End If
        ' �q�ɓ�����������
        Response.Write anyTmp(44) & ","
        ' �f�o��
        Response.Write anyTmp(45) & ","
        ' ��R���ԋp
        Response.Write anyTmp(46) & ","

        ' �����m�F�\�莞��
        Response.Write anyTmp(62) & ","
        ' �����m�F��������
        Response.Write Left(anyTmp(48),10) & ","
        ' ���A��
        If anyTmp(47)="S" Then
            Response.Write "�~,"
        ElseIf anyTmp(47)="C" Then
            Response.Write "��,"
        Else
            Response.Write "�|,"
        End If
        ' �ʔ���
        If anyTmp(63)<>"" Then
            Response.Write "��,"
        Else
            Response.Write "�~,"
        End If
        ' �ʊց^�ېŗA��
        If anyTmp(49)<>"" Then
            Response.Write "��,"
        Else
            Response.Write "�~,"
        End If
        ' �����DO���s
        If anyTmp(51)<>"Y" Then
            Response.Write "�~,"
        Else
            Response.Write "��,"
        End If
        ' �t���[�^�C��
        Response.Write Left(anyTmp(52),10) & ","
        ' ���o��
        If anyTmp(34)="Y" Then
            Response.Write "��,"
        ElseIf anyTmp(34)="S" Then
            Response.Write "��,"
        Else
            Response.Write "�~,"
        End If

        ' �T�C�Y
        If anyTmp(53)<>"" Then
            Response.Write anyTmp(53) & ","
        Else
            Response.Write anyTmp(10) & ","
        End if
        ' ����
        Response.Write anyTmp(54) & ","
        ' ���[�t�@�[
        If anyTmp(55)="R" Then
            Response.Write "��,"
        ElseIf anyTmp(55)<>"" Then
            Response.Write "�|,"
        Else
            If anyTmp(11)<>"" Then
                If anyTmp(11)<>"RF" Then
                    Response.Write "�|,"
                Else
                    Response.Write "��,"
                End If
            Else
                Response.Write ","
            End If
        End If
        ' �d��
        If anyTmp(56)<>"" And anyTmp(56)<>"0" Then
            dWeight=anyTmp(56) / 10
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
        Response.Write ","
        ' �댯��
        If anyTmp(57)="H" Then
            Response.Write "��,"
        ElseIf anyTmp(57)<>"" Then
            Response.Write "�|,"
        Else
            Response.Write ","
        End If
        ' ���o�^�[�~�i����
        Response.Write anyTmp(35) & ","
        ' �X�g�b�N���[�h���p
        If anyTmp(65)>="1" And anyTmp(65)<="4" Then
            Response.Write "��,"
        Else
            Response.Write "�~,"
        End If
        ' �ԋp�ꏊ
        Response.Write anyTmp(40) & ","

        If bMode="���^" Then
            ' �D�Ж�
            If anyTmp(36)<>"" Then
                Response.Write anyTmp(36) & ","
            ElseIf anyTmp(21)<>"" Then
                Response.Write anyTmp(21) & ","
            Else
                Response.Write anyTmp(15) & ","
            End If
            ' �D��
            If anyTmp(37)<>"" Then
                Response.Write anyTmp(37) & ","
            Else
                Response.Write anyTmp(2) & ","
            End If
            ' ���q
            If anyTmp(38)<>"" Then
                Response.Write anyTmp(38) & ","
            Else
                Response.Write anyTmp(3) & ","
            End If
        End If
        ' �d�o�`
        Response.Write anyTmp(39) & ","
        ' �O�`
        Response.Write anyTmp(68)

        ' �ېŗA��(From)
'        Response.Write anyTmp(58) & ","
        ' �ېŗA��(To)
'        Response.Write anyTmp(59)

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>