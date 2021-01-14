<%
'�A�o�Ɖ�֌W�֐�

'�A�o�R���e�i�Ɖ�
Function SerchMSExpCntnr(conn, rsd, pfp, sWhere, sOrder)
    'conn       [ I ]DB�ڑ�Object
    'rsd        [ I ]���R�[�h�Z�b�gObject
    'pfp        [ I ]�e���|�����t�@�C��
    'sWhere     [ I ]��������
    'sOrdr      [ I ]Sort����

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    bWrite = 0        '�o�̓��R�[�h����

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

    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1

    Do While Not rsd.EOF
        sText=""

        ' �o�̓��R�[�h�̕ҏW
        sText = Trim(rsd("BookNo")) & ","                    ' 1:Booking�ԍ�
        sText = sText & Trim(rsd("ContNo")) & ","            ' 2:�R���e�iNo.
        sWork = Trim(rsd("ShipFullName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("VslCode")) & ","       ' 3:�D��(Full)
        Else
            sText = sText & sWork & ","                      ' 3:�D��
        End If
        sText = sText & Trim(rsd("LdVoyage")) & ","          ' 4:���q
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
        sText = sText & Trim(rsd("ContHeight")) & ","        '13:����
        sText = sText & Trim(rsd("Remark")) & ","            '14:�R�����g
        sText = sText & Trim(rsd("ShipCtrl")) & ","          '15:�׎�Ǘ��ԍ�
        sText = sText & DispDateTime(rsd("WHArTime"),0) & ","     '16:�q�ɓ����w�����

        sWork = Trim(rsd("OpeName"))
        If sWork="" Or IsNull(sWork) Then
            sText = sText & Trim(rsd("OpeCode")) & ","       '17:�`�^����
        Else
            sText = sText & sWork & ","                      '17:�`�^����
        End If

'        sText = sText & DispDateTime(rsd("CYRecDate"),0) & ","    '17:CY�����w�����
        sText = sText & Trim(rsd("TruckRefFlag")) & ","      '18:���^�Q�ƃt���O
        sText = sText & Trim(rsd("ShipLine")) & ","          '19:�D�ЃR�[�h
        sText = sText & Trim(rsd("BookVoyage")) & ","        '20:���q(�ԍ�)
        sText = sText & Trim(rsd("PickPlace")) & ","         '21:��R�����ꏊ
'        sWork = Trim(rsd("ShipperFullName"))
'        If sWork="" Or IsNull(sWork) Then
'            sText = sText & Trim(rsd("Shipper")) & ","       '22:�׎�(�t������)
'        Else
'            sText = sText & sWork & ","                      '22:�׎�(�t������)
'        End If
'        sWork = Trim(rsd("ForFullName"))
'        If sWork="" Or IsNull(sWork) Then
'            sText = sText & Trim(rsd("Forwarder")) & ","     '23:�C��(�t������)
'        Else
'            sText = sText & sWork & ","                      '23:�C��(�t������)
'        End If
		
        sText = sText & DispDateTime(rsd("CYCut"),0) & ","    '22:�J�b�g��
        sText = sText & Trim(rsd("RecTerminal")) & ","       '23:CY�����ꏊ

'        sWork = Trim(rsd("TruckFullName"))
'       If sWork="" Or IsNull(sWork) Then
'            sText = sText & Trim(rsd("Trucker")) & ","       '24:���^(�t������)
'        Else
'            sText = sText & sWork & ","                      '24:���^(�t������)
'        End If

        sText = sText & Trim(rsd("Shipper")) & ","      	' 24:�׎�R�[�h

        sText = sText & DispDateTime(rsd("PickDate"),0) & "," '25:��R�����o�w���
        sText = sText & Trim(rsd("OpeCode")) & ","			'26:�`�^�R�[�h
        sText = sText & Trim(rsd("PickOKFlag")) & ","		'27:��R�������ۃt���O
        sText = sText & Trim(rsd("PkPlUpdFlag")) & ","		'28:�s�b�N�ꏊ�ύX�t���O
        sText = sText & Trim(rsd("PkDtUpdFlag")) & ","		'29:���o�w����ύX�t���O

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
            'SQL�𔭍s���Č���
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

            'SQL�𔭍s���Č���
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' �o�̓��R�[�h�̕ҏW
                strOut = Trim(rsd("BookNo")) & ","                     ' 1:Booking No.
                strOut = strOut & Trim(rsd("ContNo")) & ","            ' 2:�R���e�iNo.
                strOut = strOut & Trim(rsd("PickPlace")) & ","         ' 3:��R�����ꏊ
                strOut = strOut & Trim(rsd("ContSize")) & ","          ' 4:�T�C�Y
                strOut = strOut & Trim(rsd("ContHeight")) & ","        ' 5:����
                ' 6:���[�t�@�[
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
                strOut = strOut & Trim(rsd("RecTerminal")) & ","       ' 7:�����^�[�~�i����
                strOut = strOut & Trim(rsd("SealNo")) & ","            ' 8:�V�[��No.
                strOut = strOut & Trim(rsd("ContWeight")) & ","        ' 9:���d��
                strOut = strOut & DispDateTime(rsd("CYOpen"),10) & "," '10:�b�x�I�[�v��
                strOut = strOut & DispDateTime(rsd("CYCut"),10) & ","  '11:�b�x�N���[�Y

                If IsNull(rsd("ShipLineName")) Then
                    strOut = strOut & Trim(rsd("ShipLine")) & ","      '12:�D�Ж�
                Else
                    strOut = strOut & Trim(rsd("ShipLineName")) & ","  '12:�D�Ж�
                End If
                If IsNull(rsd("ShipName")) Then
                    strOut = strOut & Trim(rsd("VslCode")) & ","       '13:�D��
                Else
                    strOut = strOut & Trim(rsd("ShipName")) & ","      '13:�D��
                End If
                strOut = strOut & Trim(rsd("LdVoyage")) & ","          '14:���q
                strOut = strOut & Trim(rsd("DPort")) & ","             '15:�d���`(�R�[�h)

                ' ���n��̃`�F�b�N
                strEmpDelTime = DispDateTime(rsd("EmpDelTime"),0)      '17:��R����掞��
                strWHArTime = DispDateTime(rsd("WHArTime"),0)          '18:�q�ɓ���
                strVanTime = DispDateTime(rsd("VanTime"),0)            '19:�o���j���O
                strSYRecTime = DispDateTime(rsd("SYRecTime"),0)        '25:�X�g�b�N���[�h��������
                strExpArTime = DispDateTime(rsd("ExpArTime"),0)        '20:�b�x����
                strOutTime = DispDateTime(rsd("OutTime"),0)            '21:�D�ϊ���
                If strWHArTime<strEmpDelTime Then
                    strWHArTime = ""                    ' �q�ɓ�������R����掞�����ߋ��̂Ƃ��\�����Ȃ�
                End If
                strFuture = strSYRecTime
                If strFuture="" Then
                    strFuture = strExpArTime
                End If
                If strFuture="" Then
                    strFuture = strOutTime
                End If
                If strFuture<>"" And strWHArTime>strFuture Then
                    strWHArTime = ""                    ' �q�ɓ�����KACCS,Hits������薢���̂Ƃ��\�����Ȃ�
                End If
                strOld = strWHArTime
                If strOld="" Then
                    strOld = strEmpDelTime
                End If
                If strVanTime<strOld Then
                    strVanTime = ""                     ' �o���j���O���q�ɓ����܂��͋�R����掞�����ߋ��̂Ƃ��\�����Ȃ�
                End If
                If strFuture<>"" And strVanTime>strFuture Then
                    strVanTime = ""                     ' �o���j���O��KACCS,Hits������薢���̂Ƃ��\�����Ȃ�
                End If

                strCheck=rsd("ShipName")                '�D�����Ȃ�������A���݁E���ݎ�����\�����Ȃ��悤�ɂ���

                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & DispDateTime(rsd("ETD"),0) & ","        '16:���ݗ\�莞��(�^�[�~�i��)
                End If
                strOut = strOut & DispDateTime(rsd("EmpDelTime"),0) & ","     '17:��R����掞��
                strOut = strOut & strWHArTime & ","                           '18:�q�ɓ���
                strOut = strOut & strVanTime & ","                            '19:�o���j���O
                strOut = strOut & DispDateTime(rsd("ExpArTime"),0) & ","      '20:�b�x����
                strOut = strOut & DispDateTime(rsd("OutTime"),0) & ","        '21:�D�ϊ���
                If IsNull( strCheck ) Then
                    strOut = strOut & ",,,"
                Else
                    strOut = strOut & DispDateTime(rsd("TD"),0) & ","         '22:���݊�������(�^�[�~�i��)
                    strOut = strOut & ",,"
'                    strOut = strOut & DispDateTime(rsd("TA"),0) & ","         '23:���݊�������(�d���`)
'                    strOut = strOut & DispDateTime(rsd("ETA"),0) & ","        '24:���ݗ\�莞��(�d���`)
                End If
                strOut = strOut & DispDateTime(rsd("SYRecTime"),0) & ","      '25:�X�g�b�N���[�h��������
                If IsNull( strCheck ) Then
                    strOut = strOut & ","
                Else
                    strOut = strOut & DispDateTime(rsd("ETDLong"),0) & ","    '26:���݃X�P�W���[������(�^�[�~�i��)
                End If
'                strOut = strOut & DispDateTime(rsd("WHArSchDate"),0) & ","    '27:��q�ɓ����X�P�W���[������
                strOut = strOut & ","                                         '27:��q�ɓ����X�P�W���[������(�ʏ�e�[�u���ɂ͐ݒ肳��Ȃ�)

                strOut = strOut & Trim(rsd("CargoWeight")) & ","              '28:�ݕ��d��
                strOut = strOut & Trim(rsd("SYRecFlag")) & ","                '29:�X�g�b�N���[�h�����\��t���O
                strOut = strOut & DispDateTime(rsd("SYPRRecTime"),0) & ","    '30:�X�g�b�N���[�h�����\�����
                strOut = strOut & DispDateTime(rsd("CYRecPrDate"),0) & ","    '31:�b�x�����w�����

		        '32:�댯�i
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

				strOut = strOut & Trim(rsd("ContType"))              '33:�R���e�i�^�C�v

                ' �d���`�f�[�^��T��
                strTmp=Split(strOut,",")
                If strTmp(14)<>"" Then
                    rsd.Close

                    sql = "SELECT ExportCont.VoyCtrl, VslPort.TA, VslPort.ETA FROM ExportCont, VslPort " & _
                          "WHERE (" & sWhere & ") AND VslPort.VslCode=ExportCont.VslCode AND " & _
                          "VslPort.VoyCtrl=ExportCont.VoyCtrl AND VslPort.PortCode='" & strTmp(14) & "'"
                    'SQL�𔭍s���Č���
                    rsd.Open sql, conn, 0, 1, 1

                    If Not rsd.EOF Then
                        If strTmp(12)<>"" Then    '�D������������
                            strTmp(22) = DispDateTime(rsd("TA"),0)     '23:���݊�������(�d���`)
                            strTmp(23) = DispDateTime(rsd("ETA"),0)    '24:���ݗ\�莞��(�d���`)
                        End If
                    End If
                    rsd.Close

                    sql = "SELECT mPort.FullName From mPort " & _
                          "WHERE mPort.PortCode='" & strTmp(14) & "'"
                    'SQL�𔭍s���Č���
                    rsd.Open sql, conn, 0, 1, 1

                    If Not rsd.EOF Then
                        If Not IsNull(rsd("FullName")) Then
                            strTmp(14) = Trim(rsd("FullName"))         '15:�d���`
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

            'SQL�𔭍s���Č���
            rsd.Open sql, conn, 0, 1, 1

            If Not rsd.EOF Then
                ' �o�̓��R�[�h�̕ҏW
                strOut = Trim(rsd("BookNo")) & ","                     ' 1:Booking No.
                strOut = strOut & strTmp(1) & ","                      ' 2:�R���e�iNo.
                strOut = strOut & ","                                  ' 3:��R�����ꏊ
                strOut = strOut & ","                                  ' 4:�T�C�Y
                strOut = strOut & ","                                  ' 5:����
                strOut = strOut & Trim(rsd("RFFlag")) & ","            ' 6:���[�t�@�[
                strOut = strOut & Trim(rsd("RecTerminal")) & ","       ' 7:�����^�[�~�i����
                strOut = strOut & ","                                  ' 8:�V�[��No.
                strOut = strOut & ","                                  ' 9:���d��
                strOut = strOut & DispDateTime(rsd("CYOpen"),10) & "," '10:�b�x�I�[�v��
                strOut = strOut & DispDateTime(rsd("CYCut"),10) & ","  '11:�b�x�N���[�Y

                If IsNull(rsd("ShipLineName")) Then
                    strOut = strOut & Trim(rsd("ShipLine")) & ","      '12:�D�Ж�
                Else
                    strOut = strOut & Trim(rsd("ShipLineName")) & ","  '12:�D�Ж�
                End If
                If IsNull(rsd("ShipName")) Then
                    strOut = strOut & Trim(rsd("VslCode")) & ","       '13:�D��
                Else
                    strOut = strOut & Trim(rsd("ShipName")) & ","      '13:�D��
                End If
                strOut = strOut & Trim(rsd("LdVoyage")) & ","          '14:���q
                If IsNull(rsd("DPortName")) Then
                    strOut = strOut & Trim(rsd("DPort")) & ","         '15:�d���`
                Else
                    strOut = strOut & Trim(rsd("DPortName")) & ","     '15:�d���`
                End If

                strOut = strOut & DispDateTime(rsd("ETD"),0) & ","     '16:���ݗ\�莞��(�^�[�~�i��)
                strOut = strOut & ","                                  '17:��R����掞��
                strOut = strOut & ","                                  '18:�q�ɓ���
                strOut = strOut & ","                                  '19:�o���j���O
                strOut = strOut & ","                                  '20:�b�x����
                strOut = strOut & ","                                  '21:�D�ϊ���
                strOut = strOut & DispDateTime(rsd("TD"),0) & ","      '22:���݊�������(�^�[�~�i��)
                strOut = strOut & DispDateTime(rsd("TA"),0) & ","      '23:���݊�������(�d���`)

                strOut = strOut & DispDateTime(rsd("ETA"),0) & ","     '24:���ݗ\�莞��(�d���`)
                strOut = strOut & ","                                  '25:�X�g�b�N���[�h��������
                strOut = strOut & DispDateTime(rsd("ETDLong"),0) & "," '26:���݃X�P�W���[������(�^�[�~�i��)
                strOut = strOut & DispDateTime(rsd("WHArSchDate"),0) & ","  '27:��q�ɓ����X�P�W���[������

                strOut = strOut & ","                                  '28:�ݕ��d��
                strOut = strOut & ","                                  '29:�X�g�b�N���[�h�����\��t���O
                strOut = strOut & ","                                  '30:�X�g�b�N���[�h�����\�����
                strOut = strOut & ","                                  '31:�b�x�����w�����
                strOut = strOut & ","                                  '32:�댯�i
																	   '33:�^�C�v

            Else
                strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            rsd.Close
        Else
            strOut = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
        End If

        pfp.WriteLine Tmp(i) & strOut & ","
    Next

    SerchMSExpCntnr = bWrite    '�Y�����R�[�h���������ǂ���
End Function

'�A�o�R���e�iCSV�t�@�C���^�C�g���s�o��
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]�������[�h(���O�C�����)

    If bMode="�C��" Then
        Response.Write "�׎�,"
        Response.Write "�׎�Ǘ��ԍ�,"
    ElseIf bMode="���^" Then
        Response.Write "�C��,"
    ElseIf bMode="�׎�" Then
        Response.Write "�׎�Ǘ��ԍ�,"
    Else
        Response.Write "�׎�,"
        Response.Write "�׎�Ǘ��ԍ�,"
    End If

    Response.Write "Booking�ԍ�,"
    Response.Write "�R���e�iNo.,"

    If bMode="�C��" Then
        Response.Write "�w�藤�^�Ǝ�,"
        Response.Write "�q�ɗ���,"
    ElseIf bMode="���^" Then
        Response.Write "�׎�,"
        Response.Write "�q�ɗ���,"
    ElseIf bMode="�׎�" Then
        Response.Write "�C��,"
        Response.Write "�w�藤�^�Ǝ�,"
    Else
        Response.Write "�w�藤�^�Ǝ�,"
        Response.Write "�q�ɗ���,"
    End If

    Response.Write "��R�����ꏊ,"
    Response.Write "��R�����o�w���,"
    Response.Write "�T�C�Y,"
    Response.Write "�^�C�v,"
    Response.Write "����,"
    Response.Write "���[�t�@�[,"
    Response.Write "�`�^�R�[�h,"

    Response.Write "�V�[��No.,"
    Response.Write "�ݕ��d��,"
    Response.Write "���d��,"
    Response.Write "�����^�[�~�i����,"
    Response.Write "CY�I�[�v����,"
    Response.Write "CY�N���[�Y��,"

    Response.Write "��R����抮������,"
    Response.Write "�q�ɓ����w������,"
    Response.Write "�q�ɓ�����������,"
    Response.Write "�o���j���O��������,"
    Response.Write "SY�\�񎞍�,"
    Response.Write "SY������������,"
    Response.Write "CY�����w����,"
    Response.Write "CY������������,"
    Response.Write "�D�ϊ�������,"
    Response.Write "���݌v�掞��,"
    Response.Write "���ݗ\�莞��,"
    Response.Write "���݊�������,"
    Response.Write "�d���`���ݗ\�莞��,"
    Response.Write "�d���`���݊�������,"

    Response.Write "�D��,"
    Response.Write "�D��,"
    Response.Write "Voyage No.,"
    Response.Write "�d���`,"
    Response.Write "�댯�i"

    Response.Write Chr(13) & Chr(10)
End Sub

'�A�o�R���e�iCSV�t�@�C���f�[�^�s�o��
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]�������[�h(���O�C�����)
    'pfp        [ I ]�e���|�����t�@�C��

    ' �_�E�����[�h�t�@�C���̃��R�[�h������ԌJ��Ԃ�
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        If bMode="�C��" Then
            ' �׎喼��
            Response.Write anyTmp(21) & ","
            ' �׎�Ǘ��ԍ�
            Response.Write anyTmp(14) & ","
        ElseIf bMode="���^" Then
            ' �C��
            Response.Write anyTmp(22) & ","
        ElseIf bMode="�׎�" Then
            ' �׎�Ǘ��ԍ�
            Response.Write anyTmp(14) & ","
        Else
            ' �׎�Ǘ��ԍ�
            Response.Write anyTmp(14) & ","
        End If

        ' Booking No.
        Response.Write anyTmp(0) & ","
        ' �R���e�iNo.
        Response.Write anyTmp(1) & ","

        If bMode="�C��" Then
            ' �w�藤�^�Ǝ�
            Response.Write anyTmp(23) & ","
            ' �q�ɗ���
            Response.Write anyTmp(13) & ","
        ElseIf bMode="���^" Then
            ' �׎喼��
            Response.Write anyTmp(21) & ","
            ' �q�ɗ���
            Response.Write anyTmp(13) & ","
        ElseIf bMode="�׎�" Then
            ' �C��
            Response.Write anyTmp(22) & ","
            ' �w�藤�^�Ǝ�
            Response.Write anyTmp(23) & ","
        Else
            ' �w�藤�^�Ǝ�
            Response.Write anyTmp(23) & ","
            ' �q�ɗ���
            Response.Write anyTmp(13) & ","
        End If

        ' ��R�����ꏊ
        If anyTmp(32)<>"" Then
            Response.Write anyTmp(32) & ","
        Else
            Response.Write anyTmp(20) & ","
        End If
        ' ��R�����o�w���
        Response.Write anyTmp(24) & ","
        ' �T�C�Y
        If anyTmp(33)<>"" Then
            Response.Write anyTmp(33) & ","
        Else
            Response.Write anyTmp(10) & ","
        End If
        ' �^�C�v
        If anyTmp(62)<>"" Then
            Response.Write anyTmp(62) & ","
        Else
	        Response.Write anyTmp(11) & ","
        End If
        ' ����
        If anyTmp(34)<>"" Then
            Response.Write anyTmp(34) & ","
        Else
            Response.Write anyTmp(12) & ","
        End If
        ' ���[�t�@�[
        If anyTmp(35)="R" Then
            Response.Write "��,"
        ElseIf anyTmp(35)<>"" Then
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
        ' �`�^�R�[�h
        Response.Write anyTmp(25) & ","
        ' �V�[��No.
        Response.Write anyTmp(37) & ","
        ' �ݕ��d��
        If anyTmp(57)<>"" And anyTmp(57)<>"0" Then
            dWeight=anyTmp(57) / 10
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
        Response.Write ","
        ' ���d��
        If anyTmp(38)<>"" And anyTmp(38)<>"0" Then
            dWeight=anyTmp(38) / 10
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
        Response.Write ","
        ' �����^�[�~�i����
        Response.Write anyTmp(36) & ","
        ' �b�x�I�[�v��
        Response.Write anyTmp(39) & ","
        ' �b�x�N���[�Y
        Response.Write anyTmp(40) & ","

        ' ��R����掞��
        Response.Write anyTmp(46) & ","
        ' �q�ɓ����X�P�W���[��
        If anyTmp(56)<>"" Then
            Response.Write anyTmp(56) & ","
        Else
            Response.Write anyTmp(15) & ","
        End If
        ' �q�ɓ���
        Response.Write anyTmp(47) & ","
        ' �o���j���O
        Response.Write anyTmp(48) & ","
        ' �X�g�b�N���[�h�����\��
        sTemp=DispReserveCell(anyTmp(58),anyTmp(59),sColor)
        If anyTmp(58)>="1" And anyTmp(58)<="4" Then
            sTemp=Left(anyTmp(59),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' �X�g�b�N���[�h����
        Response.Write anyTmp(54) & ","
        ' �b�x�����w��
        Response.Write Left(anyTmp(60),10) & ","
        ' �b�x����
        Response.Write anyTmp(49) & ","
        ' �D�ϊ���
        Response.Write anyTmp(50) & ","
        ' ���݃X�P�W���[��
        Response.Write anyTmp(25) & ","
        ' ���ݗ\�莞��
        Response.Write anyTmp(45) & ","
        ' ���݊�������
        Response.Write anyTmp(51) & ","
        ' ���ݗ\�莞��
        Response.Write anyTmp(53) & ","
        ' ���݊�������
        Response.Write anyTmp(52) & ","

        ' �D�Ж�
        If anyTmp(41)<>"" Then
            Response.Write anyTmp(41) & ","
        Else
            Response.Write anyTmp(29) & ","
        End If
        ' �D��
        If anyTmp(32)<>"" Then
            Response.Write anyTmp(32) & ","
        Else
            Response.Write anyTmp(2) & ","
        End If
        ' ���q
        If anyTmp(43)<>"" Then
            Response.Write anyTmp(43) & ","
        Else
            Response.Write anyTmp(3) & ","
        End If
        ' �d���`
        Response.Write anyTmp(44) & ","

        ' �댯�i
        If anyTmp(61)="H" Then
            Response.Write "��"
        ElseIf anyTmp(61)<>"" Then
            Response.Write "�|"
        Else
            Response.Write ""
        End If

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>