<%
'�A�o�Ɖ�֌W�֐�

'�A�o�R���e�i�Ɖ�
Function SerchExpCntnr(conn, rsd, pfp, sWhere)
    'conn       [ I ]DB�ڑ�Object
    'rsd        [ I ]���R�[�h�Z�b�gObject
    'pfp        [ I ]�e���|�����t�@�C��
    'sWhere     [ I ]��������

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    bWrite = 0        '�o�̓��R�[�h����

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
'������ Add_S  by nics 2010.02.02
    sql = sql &  ", ExportCont.CustOK "
'������ Add_E  by nics 2010.02.02
'������ Add_S  by nics 2009.02.27
    sql = sql &  ", InspectCont.XRStatus, InspectCont.XDStatus "
' - NICS�s�(2009/07/22)    sql = sql &  ", (Select TOP 1 StockPlace From mPlaceCode Where mPlaceCode.Operator=Booking.RecTerminalCode) As StockPlace "
    sql = sql &  ", (Select TOP 1 StockPlace From mPlaceCode Where mPlaceCode.Operator=(Select TOP 1 mSL.OpeCode From mShipLine mSL Where mSL.ShipLine=VslSchedule.ShipLine) and mPlaceCode.Terminal=Booking.RecTerminalCode) As StockPlace "
    sql = sql &  ", (Select TOP 1 mOperator.FullName From mShipLine mSL, mOperator Where mSL.ShipLine=VslSchedule.ShipLine AND mOperator.OpeCode=mSL.OpeCode) As OpeName "
'������ Add_E  by nics 2009.02.27
    sql = sql & " FROM ExportCont, Container, Booking, ExportCargoInfo, VslSchedule, mShipLine, mVessel, VslPort "
'������ Add_S  by nics 2009.02.27
    sql = sql &  ", InspectCont "
'������ Add_E  by nics 2009.02.27
    sql = sql & "WHERE (" & sWhere & ") AND Container.VslCode=ExportCont.VslCode AND Container.VoyCtrl=ExportCont.VoyCtrl AND " & _
          "Container.ContNo=ExportCont.ContNo AND Booking.VslCode=*ExportCont.VslCode AND Booking.VoyCtrl=*ExportCont.VoyCtrl AND " & _
          "Booking.BookNo=*ExportCont.BookNo AND ExportCargoInfo.ContNo=*ExportCont.ContNo AND ExportCargoInfo.BookNo=*ExportCont.BookNo AND ExportCargoInfo.VslCode=*ExportCont.VslCode AND " & _
          "VslSchedule.VslCode=*ExportCont.VslCode AND VslSchedule.VoyCtrl=*ExportCont.VoyCtrl AND " & _
          "mShipLine.ShipLine=*Container.ShipLine AND mVessel.VslCode=*ExportCont.VslCode AND " & _
          "VslPort.VslCode=*ExportCont.VslCode AND VslPort.VoyCtrl=*ExportCont.VoyCtrl AND VslPort.PortCode='" & strPort & "'"
'������ Add_S  by nics 2009.02.27
    sql = sql &  " AND InspectCont.VslCode=*ExportCont.VslCode AND InspectCont.VoyCtrl=*ExportCont.VoyCtrl AND " & _
          "InspectCont.ContNo=*ExportCont.ContNo " 
'������ Add_E  by nics 2009.02.27
    sql = sql & " ORDER BY ExportCont.BookNo, ExportCont.ContNo"

    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1

    Dim sOutText()

    Do While Not rsd.EOF
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
        strEmpDelTime = DispDateTime(rsd("EmpDelTime"),0)             '17:��R����掞��
        strWHArTime = DispDateTime(rsd("WHArTime"),0)                 '18:�q�ɓ���
        strVanTime = DispDateTime(rsd("VanTime"),0)                   '19:�o���j���O
        strSYRecTime = DispDateTime(rsd("SYRecTime"),0)               '25:�X�g�b�N���[�h��������
        strExpArTime = DispDateTime(rsd("ExpArTime"),0)               '20:�b�x����
        strOutTime = DispDateTime(rsd("OutTime"),0)                   '21:�D�ϊ���
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

	'�D���̃`�F�b�N���~�߂�(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ","
'        Else
            strOut = strOut & DispDateTime(rsd("ETD"),0) & ","        '16:���ݗ\�莞��(�^�[�~�i��)
'        End If
	'�D���̃`�F�b�N���~�߂�(2004/03/12)
        strOut = strOut & DispDateTime(rsd("EmpDelTime"),0) & ","     '17:��R����掞��
        strOut = strOut & strWHArTime & ","                           '18:�q�ɓ���
        strOut = strOut & strVanTime & ","                            '19:�o���j���O
        strOut = strOut & DispDateTime(rsd("ExpArTime"),0) & ","      '20:�b�x����
        strOut = strOut & DispDateTime(rsd("OutTime"),0) & ","        '21:�D�ϊ���
	'�D���̃`�F�b�N���~�߂�(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ",,,"
'        Else
            strOut = strOut & DispDateTime(rsd("TD"),0) & ","         '22:���݊�������(�^�[�~�i��)
            strOut = strOut & ",,"
'            strOut = strOut & DispDateTime(rsd("TA"),0) & ","         '23:���݊�������(�d���`)
'            strOut = strOut & DispDateTime(rsd("ETA"),0) & ","        '24:���ݗ\�莞��(�d���`)
'        End If
	'�D���̃`�F�b�N���~�߂�(2004/03/12)

        strOut = strOut & DispDateTime(rsd("SYRecTime"),0) & ","      '25:�X�g�b�N���[�h��������
	'�D���̃`�F�b�N���~�߂�(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ","
'        Else
            strOut = strOut & DispDateTime(rsd("ETDLong"),0) & ","    '26:���݃X�P�W���[������(�^�[�~�i��)
'        End If
	'�D���̃`�F�b�N���~�߂�(2004/03/12)
        strOut = strOut & DispDateTime(rsd("WHArSchDate"),0) & ","    '27:��q�ɓ����X�P�W���[������

        strOut = strOut & Trim(rsd("CargoWeight")) & ","              '28:�ݕ��d��
        strOut = strOut & Trim(rsd("SYRecFlag")) & ","                '29:�X�g�b�N���[�h�����\��t���O
        strOut = strOut & DispDateTime(rsd("SYPRRecTime"),0) & ","    '30:�X�g�b�N���[�h�����\�����
        If IsNull(rsd("CYRecPrDate")) Then
            strOut = strOut & DispDateTime(rsd("CYRecDate"),0) & ","  '31:�b�x�����w�����
        Else
            strOut = strOut & DispDateTime(rsd("CYRecPrDate"),0) & ","'31:�b�x�����w�����
        End If
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
'--- mod by mes(2005/3/28)
'		strOut = strOut & strDGFlag
		strOut = strOut & strDGFlag & ","
'--- end mes

'--- add by mes(2005/3/28)
		' 33:�e�A�E�F�C�g
        strOut = strOut & Trim(rsd("TareWeight"))
'		If Trim(rsd("TareWeight"))<>"" Then
'	        strOut = strOut & (Trim(rsd("TareWeight")) * 1000)
'	    Else
'	    	strOut = strOut & Trim(rsd("TareWeight"))
'	    End If
'--- end mes

'������ Add_S  by nics 2009.02.27
        ' 34:X���L��
        If Trim(rsd("XRStatus")) <> "" then
            strOut = strOut & "," & "�L"
        Else
            strOut = strOut & "," & "��"
        End If

        ' 35:X��CY�ԋp
        Select Case Trim(rsd("XRStatus"))
            Case "S"                            ' �Ώ�
                strOut = strOut & "," & "�L"
            Case "O"                            ' ������
                strOut = strOut & "," & "���o��"
            Case "C"                            ' ����
                strOut = strOut & "," & "��"
            Case Else
                strOut = strOut & "," & "�|"
        End Select

        ' 36:�^�[�~�i�������m�F
        If Trim(rsd("ExpArTime")) <> "" then
            strOut = strOut & "," & "��"
        Else
            strOut = strOut & "," & "��"
        End If

        ' 37:���u�ꏊ�R�[�h
        strOut = strOut & "," & Trim(rsd("StockPlace"))

        ' 38:�S���I�y���[�^
        strOut = strOut & "," & Trim(rsd("OpeName"))
'������ Add_E  by nics 2009.02.27
'������ Add_S  by nics 2010.02.02
        ' 38:�A�o�ʊ֏��(�����̃R�����g���Ԉ���Ă���̂�38�̂ݏC������
        If Trim(rsd("CustOK")) = "Y" then
            strOut = strOut & "," & "��"
        Else
            strOut = strOut & "," & "��"
        End If
'������ Add_E  by nics 2010.02.02

' 39:�R���e�i�^�C�v
'Add-S MES Aoyagi 2010.11.23 �R���e�i�^�C�v��\������B
	strOut = strOut & "," & Trim(rsd("ContType"))
'Add-E MES Aoyagi 2010.11.23 �R���e�i�^�C�v��\������B

        ReDim Preserve sOutText(bWrite)
        sOutText(bWrite) = strOut
        bWrite = bWrite + 1

        rsd.MoveNext
    Loop

    rsd.Close

    For i=0 To bWrite-1
        strTmp=Split(sOutText(i),",")

        ' �d���`�f�[�^��T��
        If strTmp(14)<>"" Then
            sql = "SELECT ExportCont.VoyCtrl, VslPort.TA, VslPort.ETA FROM ExportCont, VslPort " & _
                  "WHERE ExportCont.ContNo='" & strTmp(1) & "' AND ExportCont.BookNo='" & strTmp(0) & "' AND VslPort.VslCode=ExportCont.VslCode AND " & _
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
            rsd.Close

            strWork=""
            For j=0 To UBound(strTmp)
                strWork=strWork & strTmp(j) & ","
            Next
            sOutText(i)=strWork
        End If

        pfp.WriteLine sOutText(i)
    Next

    SerchExpCntnr = bWrite    '�Y�����R�[�h���������ǂ���
End Function

'�A�o�R���e�iCSV�t�@�C���^�C�g���s�o��
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]�������[�h(true=�R���e�i���� / false=�u�b�L���O����)

    Response.Write "Booking�ԍ�,"
    Response.Write "�R���e�iNo.,"

    Response.Write "��R�����ꏊ,"
    Response.Write "�T�C�Y,"
'Add-S MES Aoyagi 2010.11.23 �R���e�i�^�C�v���\�� 
    Response.Write "�^�C�v,"
'Add-E MES Aoyagi 2010.11.23 �R���e�i�^�C�v���\��
    Response.Write "����,"
    Response.Write "���[�t�@�[,"
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
    Response.Write "�댯�i,"
'--- add by mes(2005/3/28)
	' �e�A�E�F�C�g�ǉ�
    Response.Write "TW"
'--- end mes

'������ Add_S  by nics 200902����
    Response.Write ",X�������L��"
    Response.Write ",X������CY�ԋp"
    Response.Write ",���u�ꏊ�R�[�h"
    Response.Write ",�S���I�y���[�^"
'������ Add_S  by nics 201002����
    Response.Write ",�ʊ�"
'������ Add_E  by nics 201002����
    Response.Write ",�^�[�~�i�������m�F"
'������ Add_E  by nics 200902����

    Response.Write Chr(13) & Chr(10)
End Sub

'�A�o�R���e�iCSV�t�@�C���f�[�^�s�o��
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]�������[�h(true=�R���e�i���� / false=�u�b�L���O����)
    'pfp        [ I ]�e���|�����t�@�C��

    ' �_�E�����[�h�t�@�C���̃��R�[�h������ԌJ��Ԃ�
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        ' Booking No.
        If Not bMode Then
            Response.Write anyTmp(0) & ","
        Else
            Response.Write ","
        End If
        ' �R���e�iNo.
        Response.Write anyTmp(1) & ","

        ' ��R�����ꏊ
        Response.Write anyTmp(2) & ","
        ' �T�C�Y
        Response.Write anyTmp(3) & ","
'Add-S MES Aoyagi 2010.11.23 �R���e�i�^�C�v���\��
	'�^�C�v	
	Response.Write anyTmp(39) & ","
'Add-E MES Aoyagi 2010.11.23 �R���e�i�^�C�v���\��
        ' ����
        Response.Write anyTmp(4) & ","
       ' ���[�t�@�[
        If anyTmp(5)="R" Then
            Response.Write "��,"
        ElseIf anyTmp(5)<>"" Then
            Response.Write "�|,"
        Else
            Response.Write ","
        End If
        ' �V�[��No.
        Response.Write anyTmp(7) & ","
        ' �ݕ��d��
        If anyTmp(27)<>"" And anyTmp(27)<>"0" Then
            dWeight=anyTmp(27) / 10
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
        Response.Write ","
        ' ���d��
        If anyTmp(8)<>"" And anyTmp(8)<>"0" Then
            dWeight=anyTmp(8) / 10
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
        Response.Write ","
        ' �����^�[�~�i����
        Response.Write anyTmp(6) & ","
        ' �b�x�I�[�v��
        Response.Write anyTmp(9) & ","
        ' �b�x�N���[�Y
        Response.Write anyTmp(10) & ","

        ' ��R����掞��
        Response.Write anyTmp(16) & ","
        ' �q�ɓ����X�P�W���[��
        Response.Write anyTmp(26) & ","
        ' �q�ɓ���
        Response.Write anyTmp(17) & ","
        ' �o���j���O
        Response.Write anyTmp(18) & ","
        ' �X�g�b�N���[�h�����\��
        sTemp=DispReserveCell(anyTmp(28),anyTmp(29),sColor)
        If anyTmp(28)>="1" And anyTmp(28)<="4" Then
            sTemp=Left(anyTmp(29),5) & sTemp
        End If
        If Left(sTemp,1)<>"<" Then
            Response.Write sTemp
        End If
        Response.Write ","
        ' �X�g�b�N���[�h����
        Response.Write anyTmp(24) & ","
        ' �b�x�����w��
        Response.Write anyTmp(30) & ","
        ' �b�x����
        Response.Write anyTmp(19) & ","
        ' �D�ϊ���
        Response.Write anyTmp(20) & ","
        ' ���݃X�P�W���[��
        Response.Write anyTmp(25) & ","
        ' ���ݗ\�莞��
        Response.Write anyTmp(15) & ","
        ' ���݊�������
        Response.Write anyTmp(21) & ","
        ' ���ݗ\�莞��
        Response.Write anyTmp(23) & ","
        ' ���݊�������
        Response.Write anyTmp(22) & ","

        ' �D�Ж�
        Response.Write anyTmp(11) & ","
        ' �D��
        Response.Write anyTmp(12) & ","
        ' ���q
        Response.Write anyTmp(13) & ","
        ' �d���`
        Response.Write anyTmp(14) & ","

        ' �댯�i
        If anyTmp(31)="H" Then
            Response.Write "��"
        ElseIf anyTmp(31)<>"" Then
            Response.Write "�|"
        Else
            Response.Write ""
        End If

'--- add by mes(2005/3/28)
        Response.Write ","

        ' �e�A�E�F�C�g
        If anyTmp(32)<>"" And anyTmp(32)>0 Then
	    	If anyTmp(32)<100 then
		        dWeight=anyTmp(32) * 100
		    Else
		        dWeight=anyTmp(32)
		    End If
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
'--- end mes

'������ Add_S  by nics 200902����
        'X�������L��
        Response.Write "," & anyTmp(33)
        'X������CY�ԋp
        Response.Write "," & anyTmp(34)
        '���u�ꏊ�R�[�h
        Response.Write "," & anyTmp(36)
        '�S���I�y���[�^
        Response.Write "," & anyTmp(37)
'������ Add_S  by nics 201002����
        '�Ŋ�
        Response.Write "," & anyTmp(38)
'������ Add_E  by nics 201002����
        '�^�[�~�i�������m�F
        Response.Write "," & anyTmp(35)
'������ Add_E  by nics 200902����

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>