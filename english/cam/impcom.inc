<%
'�A���Ɖ�֌W�֐�

'�A���R���e�i�Ɖ�
Function SerchImpCntnr(conn, rsd, pfp, sWhere)
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

    sql = "SELECT ImportCont.BLNo, ImportCont.ContNo, ImportCont.RFFlag, ImportCont.DGFlag, " & _
          "ImportCont.VslCode, ImportCont.VoyCtrl, ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, " & _
          "ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, ImportCont.ReturnPlace, " & _
          "ImportCont.InTime, ImportCont.CYDelTime, ImportCont.SYDelTime, ImportCont.WHArDate, ImportCont.DevanTime, " & _
          "ImportCont.ReturnTime, ImportCont.CustStatus, ImportCont.FreeTimeExt, ImportCont.WHArSchDate ImCntnrWHArSch, " & _
          "ImportCont.SYDelFlag, ImportCont.SYPRDelTime, " & _
          "ImportCont.OLTICDate, " & _
          "Container.ContSize, Container.ContType, Container.ContHeight, Container.ContWeight, Container.ListDate, Container.StackPlace, Container.ShipLine, " & _
          "Container.ListNo, Container.OffDockFlag, Container.DsListFlg, " & _
          "InspectCont.BotStatus, InspectCont.ZooStatus, BL.RecTerminal, BL.IndvRecFlag, BL.IndvRecSchTime, BL.WHArSchDate BLWHArSch, BL.LPort, " & _
          "ImportCont.RFFlag BRFFlag, BL.DGFlag BDGFlag, " & _
          "ImportCargoInfo.WHArTime, VslSchedule.DsVoyage, VslSchedule.BatchRecSchTime, " & _
          "mShipLine.FullName ShipLineName, " & _
          "mVessel.FullName ShipName, " & _
          "VslPort.ETA, VslPort.TA, VslPort.ETALong"
'������ Add_S  by nics 2009.03.05
    sql = sql &  ", InspectCont.XRStatus, InspectCont.XDStatus, ImportCont.ReturnDateEST "
    sql = sql &  ", (Select TOP 1 StockPlace From mPlaceCode Where mPlaceCode.Operator=BL.RecTerminalCode) As StockPlace "
    sql = sql &  ", (Select TOP 1 mOperator.FullName From mShipLine mSL, mOperator Where mSL.ShipLine=VslSchedule.ShipLine AND mOperator.OpeCode=mSL.OpeCode) As OpeName "
'������ Add_E  by nics 2009.03.05
    sql = sql & " FROM ImportCont, Container, InspectCont, BL, ImportCargoInfo, VslSchedule, mShipLine, mVessel, VslPort "
    sql = sql & "WHERE (" & sWhere & ") AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl AND " & _
          "Container.ContNo=ImportCont.ContNo AND InspectCont.VslCode=*ImportCont.VslCode AND InspectCont.VoyCtrl=*ImportCont.VoyCtrl AND " & _
          "InspectCont.ContNo=*ImportCont.ContNo AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl AND " & _
          "BL.BLNo=*ImportCont.BLNo AND ImportCargoInfo.ContNo=*ImportCont.ContNo AND ImportCargoInfo.VslCode=*ImportCont.VslCode AND ImportCargoInfo.BLNo=*ImportCont.BLNo AND " & _
          "VslSchedule.VslCode=*ImportCont.VslCode AND VslSchedule.VoyCtrl=*ImportCont.VoyCtrl AND " & _
          "mShipLine.ShipLine=*Container.ShipLine AND mVessel.VslCode=*ImportCont.VslCode AND " & _
          "VslPort.VslCode=*ImportCont.VslCode AND VslPort.VoyCtrl=*ImportCont.VoyCtrl AND VslPort.PortCode='" & strPort & "'"
'������ Add_S  by nics 2010.02.18
''''    sql = sql & " ORDER BY ImportCont.BLNo, ImportCont.ContNo"
    sql = sql & " ORDER BY ImportCont.UpdtTime DESC, ImportCont.BLNo, ImportCont.ContNo "
'������ Add_E  by nics 2010.02.18
    'SQL�𔭍s���Č���
    rsd.Open sql, conn, 0, 1, 1

    Dim sOutText()
'������ Add_S  by nics 2009.03.05
    Dim sInfoText()
'������ Add_E  by nics 2009.03.05
'������ Add_S  by nics 2010.02.18
	strchkContNo = ""
'������ Add_E  by nics 2010.02.18

    Do While Not rsd.EOF
'������ Add_S  by nics 2010.02.18
		If strchkContNo = "" or strchkContNo <> Trim(rsd("ContNo")) Then
'������ Add_E  by nics 2010.02.18�iEnd if��MoveNext�t�߁A�L�q���āA�l�X�g�͂��Ă��܂���j
        '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r����(2003/10/25 Add)
        strchkNow = DispDateTime( Now, 8 )
        strchkFreeTime = DispDateTime( rsd("FreeTime"), 8 )
        strchkDelPermitDate = DispDateTime( rsd("DelPermitDate"), 8 )
        '���o�\����NULL�̂Ƃ����o�\�ɂȂ�̂�h�� (2004/09/01 Add)
        If strchkDelPermitDate = "" Then
            strchkDelPermitDate = "99999999"
        End If
        strchkOLTDateFrom = DispDateTime( rsd("OLTDateFrom"), 8 )
        strchkOLTDateTo = DispDateTime( rsd("OLTDateTo"), 8 )
        strchkFreeTimeExt = DispDateTime( rsd("FreeTimeExt"), 8 )
        '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r����(2003/10/25)
'--- add by mes(2005/3/28)
        ' �I���h�b�N�ŉ����X�g�ΏۊO�t���O���ΏۊO�łȂ��Ȃ牵���X�g�ԍ��̗L�����`�F�b�N
        If Trim(rsd("OffDockFlag"))="N" And (Trim(rsd("DsListFlg"))<>"1" Or IsNull(Trim(rsd("DsListFlg"))) = True ) Then
        	If Trim(rsd("ListNo"))="" Then
        		strTsukan = ""
        	End If
        End If
'--- end mes
        ' �ʊց^�ېŗA���̃`�F�b�N
        If Trim(rsd("OLTICFlag"))="I" Then
            If Trim(rsd("OLTICNo"))<>"" Then
                strTsukan = "I"
            Else
                strTsukan = ""
            End If
		'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18 Chenge)
'       Else
		ElseIf Trim(rsd("OLTICFlag"))<>"" Then
		'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18)
            '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25 Chenge)
'            If Now>=rsd("OLTDateFrom") And Now<=rsd("OLTDateTo") Then
            If strchkNow>=strchkOLTDateFrom And strchkNow<=strchkOLTDateTo Then
            '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25)
                strTsukan = "O"
            Else
                strTsukan = ""
            End If
		'OLTICFlag���󔒂̂Ƃ��A�����Ƌ��ԍ��̃`�F�b�N�����A�ʊ�OK�Ƃ���(2005/02/03 mod)
        Else
			If DispDateTime(rsd("OLTICDate"),8)<>"" AND Trim(rsd("OLTICNo"))<>"" Then
				strTsukan = "N"
			End If
		'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18)
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

	'�D���̃`�F�b�N���~�߂�(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ",,"
'        Else
            strOut = strOut & DispDateTime(rsd("ETA"),0) & "," ' 3:���ݗ\�莞��
            strOut = strOut & DispDateTime(rsd("TA"),0) & ","  ' 4:���݊�������
'        End If
	'�D���̃`�F�b�N���~�߂�(2004/03/12)
        ' 5:���o��
        If DispDateTime(rsd("CYDelTime"),0)<>"" Then
            strOut = strOut & "S,"
        '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25 Chenge)
'        ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
'           (rsd("FreeTime")>=Now Or rsd("FreeTimeExt")>=Now) And _
'           rsd("DelPermitDate")<=Now Then
'--- mod by MES(2004/8/24)
'        ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
'           (strchkFreeTime>=strchkNow Or strchkFreeTimeExt>=strchkNow) And _
'           strchkDelPermitDate<=strchkNow Then
        ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
           (strchkFreeTime>=strchkNow Or strchkFreeTimeExt>=strchkNow) And _
           strchkDelPermitDate<=strchkNow And _
           DispDateTime(rsd("OLTICDate"),8)<>"" AND Trim(rsd("OLTICNo"))<>"" Then
'--- end MES
            '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25)
'--- mod by mes(2005/3/28)
'            strOut = strOut & "Y,"
	        ' �I���h�b�N�ŉ����X�g�ΏۊO�t���O���ΏۊO�łȂ��Ȃ牵���X�g�ԍ��̗L�����`�F�b�N
	        If Trim(rsd("OffDockFlag"))="N" And (Trim(rsd("DsListFlg"))<>"1" Or IsNull(Trim(rsd("DsListFlg"))) = True ) Then
	        	If Trim(rsd("ListNo"))="" Then
		            strOut = strOut & ","
		        Else
		        	strOut = strOut & "Y,"
	        	End If
	        Else
	        	strOut = strOut & "Y,"
	        End If
'--- end mes
        Else
            strOut = strOut & ","
        End If
'        strOut = strOut & Trim(rsd("StackPlace")) & ","        ' 6:���o�^�[�~�i����
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
'            strOut = strOut & DispDateTime(rsd("TD"),0) & ","     '12:���݊�������
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
        '28:�댯�i
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
	'�D���̃`�F�b�N���~�߂�(2004/03/12 Chenge)
'        If IsNull( strCheck ) Then
'            strOut = strOut & ","
'        Else
            strOut = strOut & DispDateTime(rsd("ETALong"),0) & ","       '32:���݃X�P�W���[��
'        End If
	'�D���̃`�F�b�N���~�߂�(2004/03/12)
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
        If strWHA="" Then
            strWHA = DispDateTime(rsd("WHArTime"),0)
        End If
        strOut = strOut & strWHA & ","
        strOut = strOut & Trim(rsd("SYDelFlag")) & ","                   '36:�X�g�b�N���[�h���o�\��t���O
        strOut = strOut & DispDateTime(rsd("SYPRDelTime"),0) & ","        '37:�X�g�b�N���[�h���o�\�񎞍�

'������ Add_S  by nics 2009.03.05
        ' 38�`41 �̏��͉��L�Őݒ肳���̂ŁA42�`�̏���ʓr�z��ɐݒ�
        strInfo = ""
        ' 42:X���L��
'������ Mod_S  by nics 2010.03.23�@�p�ꃂ�[�h�֏㏑���ύX
        If Trim(rsd("XRStatus")) <> "" or Trim(rsd("XDStatus")) <> "" then
            strInfo = strInfo & "," & "Need"		
        Else
            strInfo = strInfo & "," & "Not need"	
        End If

        ' 43:X��CY�ԋp
        Select Case Trim(rsd("XRStatus"))
            Case "S"                            ' �Ώہ@�@�@�L
                strInfo = strInfo & "," & "Will be returned"
            Case "O"                            ' ������@�@���o��
                strInfo = strInfo & "," & "CY Out(will be returned)"
            Case "C"                            ' �����@�@�@��
                strInfo = strInfo & "," & "Returned"
            Case Else				' ���@�@�@�@�|
                strInfo = strInfo & "," & "Other"
        End Select
'������ Mod_E  by nics 2010.03.23�@�p�ꃂ�[�h�֏㏑���ύX

        ' 44:���u�ꏊ�R�[�h
        strInfo = strInfo & "," & Trim(rsd("StockPlace"))

        ' 45:��o���ԋp�\���
        strInfo = strInfo & "," & DispDateTime(rsd("ReturnDateEST"),10)

        ' 46:�S���I�y���[�^
        strInfo = strInfo & "," & Trim(rsd("OpeName"))

'Add-S MES Aoyagi 2010.11.27 �R���e�i�^�C�v��\���B
	strInfo = strInfo & "," & Trim(rsd("ContType")) '46:�R���e�i�^�C�v
'Add-E MES Aoyagi 2010.11.27 �R���e�i�^�C�v��\���B

        ReDim Preserve sInfoText(bWrite)
        sInfoText(bWrite) = strInfo
'������ Add_E  by nics 2009.03.05

        ReDim Preserve sOutText(bWrite)
        sOutText(bWrite) = strOut
        bWrite = bWrite + 1

'������        by nics 2010.02.18�i�ȉ���End If�́ADo While �̎���IF���ɑ΂�����́j
		End If
		strchkContNo = Trim(rsd("ContNo"))
'������ Add_E  by nics 2010.02.18
        rsd.MoveNext
    Loop
    rsd.Close

    '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2003/09/24 Add)
    iNotWriteRecords = 0
    '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2003/09/24)
    For i=0 To bWrite-1
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2003/09/24 Add)
        If sOutText(i)<>"" Then
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2003/09/24)
        strTmp=Split(sOutText(i),",")
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2003/09/24 Add)
        '����R���e�i����������Ƃ��A��ɂ܂Ƃ߂�
        For j=i+1 To bWrite-1
            If sOutText(j)<>"" Then             '(2004/07/09)
            strTmp2=Split(sOutText(j),",")
            If strTmp(1)=strTmp2(1) Then
                sOutText(j)=""
            End If
            End If                              '(2004/07/09)
        Next

        '�R���e�i�P�ʂɓǂݒ����ă`�F�b�N����B�PBL�ł��ʊ֕t��������΁A���o��s���Ƃ���
        sOutFlag = "Y"
'--- mod by MES(2004/8/24)
'        sql = "SELECT OLTICFlag, OLTICNo, OLTDateFrom, OLTDateTo, CYDelTime, DOStatus, FreeTime, FreeTimeExt, DelPermitDate FROM ImportCont " & _
'              "WHERE ContNo='" & strTmp(1) & "'"
'--- mod by mes(2005/3/28)
'        sql = "SELECT OLTICFlag, OLTICNo, OLTICDate, OLTDateFrom, OLTDateTo, CYDelTime, DOStatus, FreeTime, FreeTimeExt, DelPermitDate FROM ImportCont " & _
'              "WHERE ContNo='" & strTmp(1) & "'"
        sql = "SELECT ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.OLTICDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, ImportCont.CYDelTime, ImportCont.DOStatus, ImportCont.FreeTime, ImportCont.FreeTimeExt, ImportCont.DelPermitDate, Container.ListNo, Container.OffDockFlag, Container.DsListFlg FROM ImportCont, Container " & _
              "WHERE ImportCont.ContNo='" & strTmp(1) & "' AND " & _
              "Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl AND Container.ContNo=ImportCont.ContNo "
'--- end mes
'--- end MES
        'SQL�𔭍s���Č���
        rsd.Open sql, conn, 0, 1, 1

        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2004/02/17 Add)
        'If Not rsd.EOF Then
        Do While Not rsd.EOF
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2004/02/17)
            '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r����(2003/10/25 Add)
            strchkNow = DispDateTime( Now, 8 )
            strchkFreeTime = DispDateTime( rsd("FreeTime"), 8 )
            strchkDelPermitDate = DispDateTime( rsd("DelPermitDate"), 8 )
            '���o�\����NULL�̂Ƃ����o�\�ɂȂ�̂�h�� (2004/09/01 Add)
            If strchkDelPermitDate = "" Then
                strchkDelPermitDate = "99999999"
            End If
            strchkOLTDateFrom = DispDateTime( rsd("OLTDateFrom"), 8 )
            strchkOLTDateTo = DispDateTime( rsd("OLTDateTo"), 8 )
            strchkFreeTimeExt = DispDateTime( rsd("FreeTimeExt"), 8 )
            '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r����(2003/10/25)
            ' �ʊց^�ېŗA���̃`�F�b�N
            If Trim(rsd("OLTICFlag"))="I" Then
                If Trim(rsd("OLTICNo"))<>"" Then
                    strTsukan = "I"
                Else
                    strTsukan = ""
                End If
			'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18 Chenge)
'           Else
			ElseIf Trim(rsd("OLTICFlag"))<>"" Then
			'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18)
                '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25 Chenge)
'                If Now>=rsd("OLTDateFrom") And Now<=rsd("OLTDateTo") Then
                If strchkNow>=strchkOLTDateFrom And strchkNow<=strchkOLTDateTo Then
                '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25)
                    strTsukan = "O"
                Else
                    strTsukan = ""
                End If
			'OLTICFlag���󔒂̂Ƃ��A�����Ƌ��ԍ��̃`�F�b�N�����A�ʊ�OK�Ƃ���(2005/02/03 mod)
            Else
				If DispDateTime(rsd("OLTICDate"),8)<>"" AND Trim(rsd("OLTICNo"))<>"" Then
					strTsukan = "N"
                End If
			'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18)
            End If
            If DispDateTime(rsd("CYDelTime"),0)<>"" Then           ' ���o����Ă����灛�Ƃ���
                If IsNull(rsd("OLTDateFrom")) Or IsNull(rsd("OLTDateTo")) Then
                    strTsukan = "S"
                Else
                    strTsukan = "T"
                End If
            End If
            ' 5:���o��
            If DispDateTime(rsd("CYDelTime"),0)<>"" Then
                sOutFlag = "S"                                      '���o��
            '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25 Chenge)
'            ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
'               (rsd("FreeTime")>=Now Or rsd("FreeTimeExt")>=Now) And _
'               rsd("DelPermitDate")<=Now Then
'--- mod by MES(2004/8/24)
'            ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
'               (strchkFreeTime>=strchkNow Or strchkFreeTimeExt>=strchkNow) And _
'               strchkDelPermitDate<=strchkNow Then
            ElseIf Trim(rsd("DOStatus"))="Y" And strTsukan<>"" And _
               (strchkFreeTime>=strchkNow Or strchkFreeTimeExt>=strchkNow) And _
               strchkDelPermitDate<=strchkNow AND _
               DispDateTime(rsd("OLTICDate"),8)<>"" AND Trim(rsd("OLTICNo"))<>"" Then
'--- end MES
                '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r�ɕύX(2003/10/25)
                '�t���O�ύX����
'--- add by mes(2005/3/28)
		        ' �I���h�b�N�ŉ����X�g�ΏۊO�t���O���ΏۊO�łȂ��Ȃ牵���X�g�ԍ��̗L�����`�F�b�N
		        If Trim(rsd("OffDockFlag"))="N" And (Trim(rsd("DsListFlg"))<>"1" Or IsNull(Trim(rsd("DsListFlg"))) = True ) Then
		        	If Trim(rsd("ListNo"))="" Then
		        		sOutFlag = ""                               '���o�s��
		                Exit Do
		        	End If
		        End If
'--- end mes
            Else
                sOutFlag = ""                                       '���o�s��
                '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2004/02/17 Add)
                Exit Do
                '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2004/02/17)
            End If
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2004/02/17 Add)
        'End If
            rsd.MoveNext
        Loop
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2004/02/17)
        rsd.Close

        If sOutFlag<>strTmp(4) Then
            '�ް��ύX���f
            strTmp(4) = sOutFlag
            strOut = ""
            For j=0 To 36
                strOut = strOut + strTmp(j) + ","
            Next
            sOutText(i) = strOut
        End If
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2003/09/24)

        ' �ύ`�f�[�^��T��
        If strTmp(9)<>"" Then
            sql = "SELECT ImportCont.VoyCtrl, VslPort.TD FROM ImportCont, VslPort " & _
                  "WHERE ImportCont.ContNo='" & strTmp(1) & "' AND ImportCont.BLNo='" & strTmp(0) & "' AND VslPort.VslCode=ImportCont.VslCode AND " & _
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
                strWork = Trim(rsd("FullName"))
                If ( Not IsNull(strWork) ) And strWork<>"" Then
                    strTmp(9) = Trim(rsd("FullName"))          '10:�d�o�`
                End If
            End If
            rsd.Close

            strWork=""
            For j=0 To UBound(strTmp)-1
                strWork=strWork & strTmp(j) & ","
            Next
            sOutText(i)=strWork
        End If

        ' �O�`�f�[�^��T��
        sql = "SELECT ImportCont.VoyCtrl, B.TD, mPort.FullName FROM ImportCont, VslPort A, VslPort B, mPort " & _
              "WHERE ImportCont.ContNo='" & strTmp(1) & "' AND ImportCont.BLNo='" & strTmp(0) & "' AND A.VslCode=ImportCont.VslCode AND " & _
              "A.VoyCtrl=ImportCont.VoyCtrl AND A.PortCode='" & strPort & "' AND " & _
              "B.VslCode=ImportCont.VslCode AND B.VoyCtrl=ImportCont.VoyCtrl AND B.CallSeq=(A.CallSeq-1) AND " & _
              "mPort.PortCode=*B.PortCode"
        'SQL�𔭍s���Č���
        rsd.Open sql, conn, 0, 1, 1

        If Not rsd.EOF Then
            If strTmp(7)<>"" Then    '�D������������
                sOutText(i) = sOutText(i) & DispDateTime(rsd("TD"),0) & ","    '38:�O�`���݊���
            Else
                sOutText(i) = sOutText(i) & ","
            End If
            sOutText(i) = sOutText(i) & Trim(rsd("FullName"))                  '39:�O�`
        Else
            sOutText(i) = sOutText(i) & ","
        End If
        rsd.Close
'HiTS ver2 ADD START by SEIKO n.Ooshige 2003/06/26
'	'�f�B�e���V�����t���[�^�C���Ǝ��O���͍�Ɣԍ��̎擾
'	sql="SELECT IPC.DetentionFreeTime, ITC1.WkNo AS WkNo1, ITC2.WkNo AS WkNo2, ITC3.WkNo AS WkNo3 "&_
'            "FROM (((ImportCont AS IPC LEFT JOIN hITCommonInfo AS ITC1 ON IPC.ContNo = ITC1.ContNo) "&_
'            "LEFT JOIN hITFullOutSelect AS ITF ON IPC.ContNo = ITF.ContNo) "&_
'            "LEFT JOIN hITCommonInfo AS ITC2 ON ITF.WkContrlNo = ITC2.WkContrlNo) "&_
'            "LEFT JOIN hITCommonInfo AS ITC3 ON IPC.BLNo = ITC3.BLNo "&_
'            "WHERE IPC.ContNo='"& strTmp(1) &"'"&_
'            "ORDER BY ITC1.UpdtTime DESC"									'ADD 2005/10/07 �ŐV�̍�Ɣԍ����o��
'        rsd.Open sql, conn, 0, 1, 1
'        sOutText(i) = sOutText(i)&","& Trim(rsd("DetentionFreeTime"))		'40:�f�B�e���V�����t���[�^�C��
'        If Not IsNull(rsd("WkNo1")) Then
'            sOutText(i) = sOutText(i)&","& Trim(rsd("WkNo1"))			'41:���O���͍�Ɣԍ�
'        ElseIf Not IsNull(rsd("WkNo2")) Then
'            sOutText(i) = sOutText(i)&","& Trim(rsd("WkNo2"))			'41:���O���͍�Ɣԍ�
'        Else
'            sOutText(i) = sOutText(i)&","& Trim(rsd("WkNo3"))			'41:���O���͍�Ɣԍ�
'        End If
'        rsd.Close
'HiTS ver2 ADD END   by SEIKO n.Ooshige 2003/06/26

'������ Add_S  by nics 2009.03.05
        ' 42�`�̏���t��
        sOutText(i) = sOutText(i) & sInfoText(i)
'������ Add_E  by nics 2009.03.05

        pfp.WriteLine sOutText(i)
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2003/09/24 Add)
        Else
            iNotWriteRecords = iNotWriteRecords + 1
        End If
        '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2003/09/24)
    Next

    '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l������(2003/09/24 Add)
''    SerchImpCntnr = bWrite    '�Y�����R�[�h���������ǂ���
    SerchImpCntnr = bWrite - iNotWriteRecords  '�Y�����R�[�h���������ǂ���
    '�A���R���e�i���o�ۂ̃`�F�b�N�ɂ����āA����BL���l�����邱���܂�(2003/09/24)
End Function

'�A���R���e�iCSV�t�@�C���^�C�g���s�o��
Sub CsvTitleWrite(bMode)
    'bMode      [ I ]�������[�h(true=�R���e�i���� / false=BL����)
'������ Add_S  by nics 201002����
'�f�[�^�擾�i���O�����́j���ǂ���
	 dim USER
	 USER       = Session.Contents("userid")
'������ Add_E  by nics 201002����

    Response.Write "BL_No.,"
    Response.Write "Container_No. ,"
    Response.Write "Load_port_DepartureTime(Actual) ,"
    Response.Write "Previous_port_DepartureTime(Actual) ,"
    Response.Write "CY_ArrivalTime(Estimate) ,"
    Response.Write "CY_ArrivalTime(Intended) ,"
    Response.Write "CY_ArrivalTime(Actual) ,"
    Response.Write "CY_InTime ,"
    Response.Write "CY_OutTime ,"
    Response.Write "CYin_ReservationTime ,"
    Response.Write "CYin_Reservation _ ,"
    Response.Write "Warehouse_ArrivalTime(Instruction) ,"
    Response.Write "Warehouse_ArrivalTime(Actual) ,"
    Response.Write "DeVanning_Time ,"
    Response.Write "EmptyContainer_ReturnTime ,"

    Response.Write "ConfirmationTimeOfArrival(Intended) ,"
    Response.Write "ConfirmationTimeOfArrival(Actual) ,"
    Response.Write "Inspection ,"
    Response.Write "Individual_Confirmation ,"
    Response.Write "CustomsClearance_BondTransportation ,"
    Response.Write "DO_Issue ,"
'������ Add_S  by nics 201002����
'�a�k�ԍ��w��܂��͎��O�����͂���̌Ăяo�����̂ݕ\��
'    Response.Write "Demurrage_FreeTime ,"
	If Not bMode  or USER <> "" Then
	    Response.Write "Demurrage_FreeTime ,"
	End If
'������ Add_E  by nics 201002����
    Response.Write "Delivery_permission ,"

    Response.Write "Size ,"
'Add-S MES Aoyagi 2010.11.27 �R���e�i�^�C�v�ǉ�
    Response.Write "Type ,"
'Add-E MES Aoyagi 2010.11.27 �R���e�i�^�C�v�ǉ�
    Response.Write "Height ,"
    Response.Write "Reefer ,"
    Response.Write "GW ,"
    Response.Write "Hazard ,"
    Response.Write "Delivery_yard ,"
    Response.Write "Use_StockYard ,"
    Response.Write "Return_place ,"

    Response.Write "Ship's_Line ,"
    Response.Write "Vessel_Name ,"
    Response.Write "Voyage No.,"
    Response.Write "Load_Port ,"
    Response.Write "Previous_Port "

'    Response.Write "�ېŗA��(From),"
'    Response.Write "�ېŗA��(To)"

'HiTS ver2 ADD START by SEIKO n.Ooshige 2003/07/09
    Response.Write ",Remark "
'������ Del_S  by nics 200902����
'    Response.Write ",WorkNo. "
'������ Del_E  by nics 200902����
'HiTS ver2 ADD END by SEIKO n.Ooshige 2003/07/09

'������ Add_S  by nics 200902����
    Response.Write ",DeliveryYard(code)"
    Response.Write ",Operater"
'������ Add_E  by nics 200902����

    Response.Write Chr(13) & Chr(10)
End Sub

'�A���R���e�iCSV�t�@�C���f�[�^�s�o��
Sub CsvDataWrite(bMode, pfp)
    'bMode      [ I ]�������[�h(true=�R���e�i���� / false=BL����)
    'pfp        [ I ]�e���|�����t�@�C��
'������ Add_S  by nics 201002����
'�f�[�^�擾�i���O�����́j���ǂ���
	dim USER
	USER       = Session.Contents("userid")
'������ Add_E  by nics 201002����

    ' �_�E�����[�h�t�@�C���̃��R�[�h������ԌJ��Ԃ�
    Do While Not pfp.AtEndOfStream
        anyTmp=Split(pfp.ReadLine,",")

        ' BL�ԍ�
        If Not bMode Then
            Response.Write anyTmp(0) & ","
        Else
            Response.Write ","
        End If
        ' �R���e�iNo.
        Response.Write anyTmp(1) & ","

        ' ���݊�������(�d�o�`)
        Response.Write anyTmp(11) & ","
        ' ���݊�������(�O�`)
        Response.Write anyTmp(37) & ","
        ' ���`�X�P�W���[��
        Response.Write Left(anyTmp(31),10) & ","
        ' ���`�\�莞��
        Response.Write anyTmp(2) & ","
        ' ���`��������
        Response.Write anyTmp(3) & ","
        ' ���[�h����
        Response.Write anyTmp(12) & ","
        ' ���[�h���o
        Response.Write anyTmp(13) & ","
        ' �X�g�b�N���[�h���o�\��
'2006/03/23
'        sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
'        If anyTmp(35)>="1" And anyTmp(35)<="4" Then
'            sTemp=Left(anyTmp(36),5) & sTemp
'        End If
'        If Left(sTemp,1)<>"<" Then
'            Response.Write sTemp
'        End If
        Response.Write ","
        ' �X�g�b�N���[�h���o
'2006/03/23
'        Response.Write anyTmp(30) & ","
        Response.Write ","
        ' �q�ɓ����X�P�W���[��
        Response.Write anyTmp(34) & ","
        ' �q�ɓ�����������
        Response.Write anyTmp(14) & ","
        ' �f�o��
        Response.Write anyTmp(15) & ","
        ' ��R���ԋp
        Response.Write anyTmp(16) & ","

        ' �����m�F�\�莞��
        Response.Write anyTmp(32) & ","
        ' �����m�F��������
        Response.Write Left(anyTmp(18),10) & ","
        ' ���A��
        If anyTmp(17)="S" Then
            Response.Write "�~,"
        ElseIf anyTmp(17)="C" Then
            Response.Write "��,"
        Else
            Response.Write "�|,"
        End If
        ' �ʔ���
        If anyTmp(33)<>"" Then
            Response.Write "��,"
        Else
            Response.Write "�~,"
        End If
        ' �ʊց^�ېŗA��
        If anyTmp(19)<>"" Then
            Response.Write "��,"
        Else
            Response.Write "�~,"
        End If
        ' �����DO���s
        If anyTmp(21)<>"Y" Then
            Response.Write "�~,"
        Else
            Response.Write "��,"
        End If
'������ Add_S  by nics 201002����
'�a�k�ԍ��w��܂��͎��O�����͂���̌Ăяo�����̂ݕ\��
        ' �t���[�^�C��
'        Response.Write Left(anyTmp(22),10) & ","
		If Not bMode  or USER <> "" Then
        	Response.Write Left(anyTmp(22),10) & ","
		End If
'������ Add_E  by nics 201002����
        ' ���o��
'������ Add_S  by nics 201003����(�p��W�L�֏㏑���ύX)
        If anyTmp(4)="Y" Then			
            Response.Write "Permitted,"
        ElseIf anyTmp(4)="S" Then		
            Response.Write "Delivered,"
        Else					
            Response.Write "Stopped,"
        End If
'������ Add_E  by nics 201003����(�p��W�L�֏㏑���ύX)

        ' �T�C�Y
        Response.Write anyTmp(23) & ","

'Add-S MES Aoyagi 2010.11.27 �R���e�i�^�C�v���\��
        ' �T�C�Y
        Response.Write anyTmp(44) & ","
'Add-E MES Aoyagi 2010.11.27 �R���e�i�^�C�v���\��
        ' ����
        Response.Write anyTmp(24) & ","
        ' ���[�t�@�[
        If anyTmp(25)="R" Then
            Response.Write "��,"
        ElseIf anyTmp(25)<>"" Then
            Response.Write "�|,"
        Else
            Response.Write ","
        End If
        ' �d��
        If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
			dWeight=anyTmp(26)
			if anyTmp(26)<1000 then
	            dWeight=anyTmp(26) / 10
			else
	            dWeight=anyTmp(26) / 1000
			end if
            Response.Write dWeight
        Else
            Response.Write "�|"
        End If
        Response.Write ","
        ' �댯�i
        If anyTmp(27)="H" Then
            Response.Write "��,"
        ElseIf anyTmp(27)<>"" Then
            Response.Write "�|,"
        Else
            Response.Write ","
        End If
        ' ���o�^�[�~�i����
'        Response.Write anyTmp(5) & ","
'������ Mod_S  by nics 200902����
'        Response.Write "�����ӓ�,"
'������
        Response.Write "�����ӓ��b�x,"
'������ Mod_E  by nics 200902����
        ' �X�g�b�N���[�h���p
'2006/03/23
'        If anyTmp(35)>="1" And anyTmp(35)<="4" Then
'            Response.Write "��,"
'        Else
            Response.Write "�~,"
'        End If
        ' �ԋp�ꏊ
        Response.Write anyTmp(10) & ","

        ' �D�Ж�
        Response.Write anyTmp(6) & ","
        ' �D��
        Response.Write anyTmp(7) & ","
        ' ���q
        Response.Write anyTmp(8) & ","
        ' �d�o�`
        Response.Write anyTmp(9) & ","
        ' �O�`
        Response.Write anyTmp(38)

        ' �ېŗA��(From)
'        Response.Write anyTmp(28) & ","
        ' �ېŗA��(To)
'        Response.Write anyTmp(29)

'HiTS ver2 ADD START by SEIKO n.Ooshige 2003/07/09
	'���}�[�N
'        Response.Write "," & anyTmp(39)
		Response.Write ","
	    If anyTmp(35)="4" Then
		    Response.Write "Cargo delivered in the next morning"
	    End If
'	'���O���͍�Ɣԍ�
'        Response.Write "," & anyTmp(40)
'HiTS ver2 ADD END by SEIKO n.Ooshige 2003/07/09

'������ Add_S  by nics 200902����
        '���u�ꏊ�R�[�h
        Response.Write "," & "6TK43"
        '�S���I�y���[�^
        Response.Write "," & "�W�F�l�b�N"
'������ Add_E  by nics 200902����

        Response.Write Chr(13) & Chr(10)
    Loop
End Sub

%>