<% 

'�\��X�V����	2003/08/27(ICCT �F��ID�ɒǉ�)
Function UpdateApp(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType(), _
			sContNoRec(), sBKNo(), sContSizeRec(), bChkA(), bChkB(), bChkC(),  _
			sContNoDel(), sChID(), sBLNo(), sContSizeDel(), sDeliverTo(), sReceiveFrom(),  _
			sTerm_CD, sNinID(), sErrMsg(), sOpeNoRec(), sOpeNoDel()) 
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�\�񎞊ԑсiHH�j
'	sRDType			[ I ]��ށiDUAL,REC,DEL,�󔒁j
'	sContNoRec 		[ I ]�����R���e�i�ԍ�
'	sBKNo			[ I ]�����u�b�L���O�ԍ�
'	sContSizeRec	[ I ]�����R���e�i�T�C�Y
'	bChkA			[ I ]�����V���[�V�ɔ��o���ڂ��Ȃ�
'	bChkB			[ I ]�����V���[�V��[�ςɎg�p
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sContNoDel 		[ I ]���o�R���e�i�ԍ�
'	sChID			[ I ]���o�V���[�VID
'	sBLNo 			[ I ]���o�a�k�ԍ�
'	sContSizeDel	[I/O]���o�R���e�i�T�C�Y
'	sDeliverTo		[ I ]���o��
'	sReceiveFrom	[ I ]������
'	sTerm_CD		[ I ]�^�[�~�i��
'	sNinID			[ I ]�F�؂h�c
'	sErrMsg			[ O ]�G���[���b�Z�[�W(0:ALL,1�`4:�e�\��)�i�����̏ꍇ�͋󔒁j
'	sOpeNoRec		[ O ]�����\��ԍ�
'	sOpeNoDel		[ O ]���o�\��ԍ�

	Dim i, sWk, iAssignCnt, bOk, iDualOpeNo, iOpeOrder, iCnt
	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime
	Dim sVslCode, sVoyage, sOpenFlag
	Dim sOpeRec, sOpeDel

	sErrMsg(0) = ""

'I20040206 S
    Dim sAn,ilen
    sAn = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'I20040206 E

	'*** �\����A�\�񎞊ԑт��L�����`�F�b�N ***
	'���ݓ��t�擾
	sDateNow = GetYMDStr(Date())

	If sDateNow > sYMD Then
		sErrMsg(0) = "��Ɠ����K�؂ł͂���܂���"
		Exit Function
	End If
	If sDateNow = sYMD Then	'���ݓ��w��̏ꍇ
		'�^�s�󋵂��擾
		Call GetOpeStatus(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime)
		sWk = TermNumToSym(iNextTime)
		If sHH < sWk or (sHH = sWk and iNextStat = 2) Then
			sErrMsg(0) = "�w�莞�ԑт̗\��͒��ߐ؂��Ă��܂�"
			Exit Function
		End If
		If sHH <> TermNumToSym(iNextTime) Then
			iOpenSlot = 0	'�J���g��
		End If
	Else
		iOpenSlot = 0	'�J���g��
	End If

	'���ԑюw��̏ꍇ�ɂ̓O���[�v�̊��蓖�Ęg���擾
	iAssignCnt = -1
	sOpenFlag = ""
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'�O���[�v�̊��蓖�Ęg���擾
		iAssignCnt = GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)

		'�O���[�v���ԑя��̎擾
		sOpenFlag = GetGrpSlotHH(conn, rsd, sGrpID, sYMD, sHH)

	End If

	'*** �\��P�`�S��o�^ ***
	For i = 0 To 3

		If sRDType(i) <> "" Then

			'�\��`�F�b�N 2003/08/27 �F��ID�ǉ�
			bOk = AppUpdCheck(conn, rsd, _
					sUsrID, sGrpID, _
					sYMD, sHH, _
					sRDType(i), _
					sContNoRec(i), sBKNo(i), sContSizeRec(i), bChkA(i), bChkB(i), bChkC(i),  _
					sContNoDel(i), sChID(i), sBLNo(i), sVslCode, sVoyage, sContSizeDel(i), _
					sOpeRec, sOpeDel, sTerm_CD, sNinID(i), _
					iAssignCnt, iOpenSlot, _
					iDualOpeNo, iOpeOrder, iCnt, _
					sErrMsg(i + 1)) 

			If bOk Then
				'�\�񏈗�
				Call UpdateApp2(conn, rsd, _
					sUsrID, sGrpID, _
					sYMD, sHH, _
					iAssignCnt, sOpenFlag, _
					iDualOpeNo, iOpeOrder, iCnt, _
					sRDType(i), _
					sContNoRec(i), sBKNo(i), sContSizeRec(i), bChkA(i), bChkB(i), bChkC(i),  _
					sContNoDel(i), sChID(i), sBLNo(i), sVslCode, sVoyage, sContSizeDel(i), sDeliverTo(i), _
					sOpeRec, sOpeDel, sReceiveFrom(i), sTerm_CD, _
					sErrMsg(i + 1), sOpeNoRec(i), sOpeNoDel(i)) 
			End If

'I20040206 S
			'�����R���e�i�ԍ��̉p�����`�F�b�N
		    For ilen = 1 To len(sContNoRec(i))
		        If InStrRev(sAn, Mid(sContNoRec(i), ilen, 1)) < 1 Then
					sErrMsg(i + 1) = "�R���e�i�ԍ��ɉp�����ȊO�̕������܂܂�Ă��܂�"
		        End If
		    Next
'I20040206 E


		End If
	Next

End Function

'�\��X�V�����Q
Function UpdateApp2(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			iAssignCnt, sOpenFlag, _
			iDualOpeNo, iOpeOrder, iCnt, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC, _
			sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, sDeliverTo, _
			sOpeRec, sOpeDel, sReceiveFrom, sTerm_CD, _
			sErrMsg, sOpeNoRec, sOpeNoDel) 
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	iAssignCnt 		[ I ]�O���[�v������
'	sOpenFlag		[ I ]�J���t���O
'	iDualOpeNo		[ I ]�f���A����Ɣԍ�
'	iOpeOrder		[ I ]��Ə���
'	iCnt			[ I ]�\��g��
'	sRDType			[ I ]��ށiDUAL,REC,DEL,�󔒁j
'	sContNoRec 		[ I ]�����R���e�i�ԍ�
'	sBKNo			[ I ]�����u�b�L���O�ԍ�
'	sContSizeRec		[ I ]�����R���e�i�T�C�Y
'	bChkA			[ I ]�����V���[�V�ɔ��o���ڂ��Ȃ�
'	bChkB			[ I ]�����V���[�V��[�ςɎg�p
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sContNoDel 		[ I ]���o�R���e�i�ԍ�
'	sChID			[ I ]���o�V���[�VID
'	sBLNo 			[ I ]���o�a�k�ԍ�
'	sVslCode		[ I ]���o�D���i�a�k�w�莞�j
'	sVoyage			[ I ]���o���q�i�a�k�w�莞�j
'	sContSizeDel		[ I ]���o�R���e�i�T�C�Y
'	sDeliverTo		[ I ]���o��
'	sOpeRec			[ I ]�����I�y���[�^
'	sOpeDel			[ I ]���o�I�y���[�^
'	sReceiveFrom		[ I ]������
'	sTerm_CD		[ I ]�^�[�~�i��
'	sErrMsg			[ O ]�G���[���b�Z�[�W�i�����̏ꍇ�͋󔒁j
'	sOpeNoRec		[ O ]�����\��ԍ�
'	sOpeNoDel		[ O ]���o�\��ԍ�


	Dim sql, bTermFlag

'I20040209 S
	Dim sYLocExt
	sYLocExt = ""
'I20040209 E

	sErrMsg = "�c�a�X�V�G���[���������܂���"

	sOpeNoRec = "0"
	sOpeNoDel = "0"

	If sRDType = "REC" or sRDType = "DUAL" Then
		'��Ɣԍ��̔���
		Call GetNewOpeNo(conn, rsd, sYMD, sOpeNoRec)
	End If
	If sRDType = "DEL" or sRDType = "DUAL" Then
		'��Ɣԍ��̔���
		Call GetNewOpeNo(conn, rsd, sYMD, sOpeNoDel)
'I20040209 S
		Call GetOutCont(conn, rsd, sContNoDel)
		If rsd.EOF = False Then
			sYLocExt = rsd("YLocExt")
		End If
		rsd.Close
'I20040209 E
	End If

	'���ԑюw��̏ꍇ
	bTermFlag = False
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		bTermFlag = True
		'�J���g���g�p����ꍇ�͂�������Z
		If iDualOpeNo = 0 and (sOpenFlag = "Y" or iAssignCnt <= iCnt) Then
			'�J���g�̌��Z
			Call SubmitOpenSlot(conn, rsd)
		End If
	End If

	'�\����񌟍�
	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iDualOpeNo)
	rsd.Open sql, conn, 0, 2, 1

	'���̐\����Ƃƃf���A���ɂȂ�ꍇ
	If iDualOpeNo > 0 Then
		rsd("UpdtTime") 	= Now()				'�X�V����
		rsd("UpdtPgCd")		= "SYWB012"			'�X�V��۸��і�
		rsd("UpdtTmnl") 	= "WEB"				'�X�V�[����
		If sOpeNoRec <> "0" Then
			rsd("DualOpeNo") = CLng(sOpeNoRec)	'�f���A����Ɣԍ�
		Else
			rsd("DualOpeNo") = CLng(sOpeNoDel)	'�f���A����Ɣԍ�
		End If
		rsd("SendFlag")    	= "Y"				'�ύX���M�׸
	End If

	If sRDType = "REC" or sRDType = "DUAL" Then
		rsd.AddNew
		rsd("OpeNo")       = sOpeNoRec			'��Ɣԍ�
		rsd("UpdtTime")    = Now()				'�X�V����
		rsd("UpdtPgCd")    = "SYWB012"			'�X�V��۸��і�
		rsd("UpdtTmnl")    = "WEB"				'�X�V�[����
		rsd("UserID")      = sUsrID				'���[�U���
		rsd("GroupID")     = sGrpID				'��ٰ�ߺ��
		rsd("AppDate")     = Now()				'�\����
		rsd("WorkDate")    = ChgYMDDate(sYMD)	'��Ɠ�
		rsd("RecDel")      = "R"				'��Ǝ��
		rsd("ContNo")      = sContNoRec			'���Ŕԍ�
		rsd("BookNo")      = sBKNo				'�ޯ�ݸޔԍ�
		rsd("Status")      = "02"				'���
		rsd("Place")       = ""					'�ꏊ
		rsd("ChassisId")   = ""					'����ID
		rsd("DelFlag")     = ""					'�폜�׸
		rsd("CReason")     = ""					'�L�����Z�����R
		rsd("WorkFlag")    = ""					'���ٍ�ƒ��׸
		rsd("SendFlag")    = "Y"				'�ύX���M�׸
		If sContSizeRec = "20" Then
			rsd("Size20Flag")  = "Y"			'20�t�B�[�g�t���O
		Else
			rsd("Size20Flag")  = ""				'20�t�B�[�g�t���O
		End If
'���������D���E���q���Z�b�g����(02/02/09)
'''		rsd("VslCode")     = ""					'�D�����
'''		rsd("Voyage")      = ""					'���q
		rsd("VslCode")     = sVslCode				'�D�����
		rsd("Voyage")      = sVoyage				'���q
		rsd("BLNo")        = ""					'�a�k�ԍ�
		rsd("Term")        = sHH				'���ԑ�
		If iDualOpeNo > 0 Then	'���̍�Ƃƃf���A���ɂȂ�ꍇ
			rsd("OpeOrder") = iOpeOrder			'��Ə���
		Else
			rsd("OpeOrder") = iOpeOrder			'��Ə���
		End If
		rsd("HopeTerm")    = sHH				'��]���ԑ�
		If iDualOpeNo > 0 Then	'���̍�Ƃƃf���A���ɂȂ�ꍇ
			rsd("DualOpeNo") = iDualOpeNo		'�f���A����Ɣԍ�
		Else
'			If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
'				rsd("DualOpeNo") = CLng(sOpeNoDel)	'�f���A����Ɣԍ�
'			Else
'				rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
'			End If
			rsd("DualOpeNo") = CLng(sOpeNoDel)	'�f���A����Ɣԍ�
		End If
		rsd("ContSize")    = sContSizeRec		'�R���e�i�T�C�Y
		If bChkA Then
			rsd("NotDelFlag") = "Y"				'���o���ڂ��Ȃ�
		Else
			rsd("NotDelFlag") = ""				'���o���ڂ��Ȃ�
		End If
		If bChkB Then
			rsd("NightFlag") = "Y"				'�[�σV���[�V
		Else
			rsd("NightFlag") = ""				'�[�σV���[�V
		End If
		If bChkC Then
			rsd("MixSizeFlag") = "Y"			'20/40���p�V���[�V
		Else
			rsd("MixSizeFlag") = ""				'20/40���p�V���[�V
		End If
		rsd("DeliverTo")   = ""					'���o��
		rsd("LockFlag")    = ""					'���b�N�t���O
		rsd("Operator")    = sOpeRec			'�I�y���[�^�R�[�h
		rsd("OpeBookNo")   = ""					'���Ѓu�b�L���O�ԍ�
		rsd("ReservNo")    = 0					'�\��ԍ�

		rsd("RecContSize") = sContSizeRec		'�������Ż���
		rsd("ContHeight")  = ""					'��������ʲ�
		rsd("ContType")    = ""					'������������
		rsd("Material")    = ""					'�������ōގ�
		rsd("TareWeight")  = 0					'��������ñ����
		rsd("Gweight")     = 0					'�������Ÿ�۽����
		rsd("SoundUnsound")  = ""				'����Sound/Unsound
		rsd("Trucker")  = ""					'�����^���Ǝ�
		rsd("ReceiveFrom")  = sReceiveFrom		'������
		rsd("CustOK")  = ""						'�������Œʊ�/���ʊ�
		rsd("Terminal")  = sTerm_CD				'�Y���^�[�~�i��(VP�\��)�b��C��
		rsd("VPBookNo")      = ""				'���o�ޯ�ݸޔԍ�(VP�\��)
		rsd("LineCode")      = ""				'�q�H�R�[�h(VP�\��)
		rsd("HeadID")        = ""				'�w�b�h�h�c(ICCT�Ή�)03/07/31
		rsd("YLocExt")       = sYLocExt			'�O�����W I20040209
	End If
	If sRDType = "DEL" or sRDType = "DUAL" Then
		rsd.AddNew
		rsd("OpeNo")       = sOpeNoDel			'��Ɣԍ�
		rsd("UpdtTime")    = Now()				'�X�V����
		rsd("UpdtPgCd")    = "SYWB012"			'�X�V��۸��і�
		rsd("UpdtTmnl")    = "WEB"				'�X�V�[����
		rsd("UserID")      = sUsrID				'���[�U���
		rsd("GroupID")     = sGrpID				'��ٰ�ߺ��
		rsd("AppDate")     = Now()				'�\����
		rsd("WorkDate")    = ChgYMDDate(sYMD)	'��Ɠ�
		rsd("RecDel")      = "D"				'��Ǝ��
		rsd("ContNo")      = sContNoDel			'���Ŕԍ�
		rsd("BookNo")      = ""					'�ޯ�ݸޔԍ�
		rsd("Status")      = "02"				'���
		rsd("Place")       = "CY"				'�ꏊ
		rsd("ChassisId")   = sChID				'����ID
		rsd("DelFlag")     = ""					'�폜�׸
		rsd("CReason")    = ""					'�L�����Z�����R
		rsd("WorkFlag")    = ""					'���ٍ�ƒ��׸
		rsd("SendFlag")    = "Y"				'�ύX���M�׸
		If sContSizeDel = "20" Then
			rsd("Size20Flag")  = "Y"			'20�t�B�[�g�t���O
		Else
			rsd("Size20Flag")  = ""				'20�t�B�[�g�t���O
		End If
		rsd("VslCode")     = sVslCode			'�D�����
		rsd("Voyage")      = sVoyage			'���q
		rsd("BLNo")        = sBLNo				'�a�k�ԍ�
		rsd("Term")        = sHH				'���ԑ�
		If Not bTermFlag Then		'���ԑюw��łȂ��ꍇ
			rsd("OpeOrder") = 0					'��Ə���
		ElseIf iDualOpeNo > 0 Then	'���̍�Ƃƃf���A���ɂȂ�ꍇ
			rsd("OpeOrder") = iOpeOrder			'��Ə���
		Else
			rsd("OpeOrder") = iOpeOrder			'��Ə���
		End If
		rsd("HopeTerm")    = sHH				'��]���ԑ�
		If iDualOpeNo > 0 Then	'���̍�Ƃƃf���A���ɂȂ�ꍇ
			rsd("DualOpeNo") = iDualOpeNo		'�f���A����Ɣԍ�
		Else
'			If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
'				rsd("DualOpeNo") = CLng(sOpeNoRec)	'�f���A����Ɣԍ�
'			Else
'				rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
'			End If
			rsd("DualOpeNo") = CLng(sOpeNoRec)	'�f���A����Ɣԍ�
		End If
		rsd("ContSize")    = sContSizeDel		'�R���e�i�T�C�Y
		rsd("NotDelFlag")  = ""					'���o���ڂ��Ȃ�
		rsd("NightFlag")   = ""					'�[�σV���[�V
		rsd("MixSizeFlag") = ""					'20/40���p�V���[�V
		rsd("DeliverTo")   = sDeliverTo			'���o��
		rsd("LockFlag")    = ""					'���b�N�t���O
		rsd("Operator")    = sOpeDel			'�I�y���[�^�R�[�h
		rsd("OpeBookNo")   = ""					'���Ѓu�b�L���O�ԍ�
		rsd("ReservNo")    = 0					'�\��ԍ�

		rsd("RecContSize") = sContSizeDel		'�������Ż���
		rsd("ContHeight")  = ""					'��������ʲ�
		rsd("ContType")    = ""					'������������
		rsd("Material")    = ""					'�������ōގ�
		rsd("TareWeight")  = 0					'��������ñ����
		rsd("Gweight")     = 0					'�������Ÿ�۽����
		rsd("SoundUnsound")  = ""				'����Sound/Unsound
		rsd("Trucker")  = ""					'�����^���Ǝ�
		rsd("ReceiveFrom")  = ""				'������
		rsd("CustOK")  = ""						'�������Œʊ�/���ʊ�
		rsd("Terminal")  = sTerm_CD				'�Y���^�[�~�i��(VP�\��)
		rsd("VPBookNo")      = ""				'���o�ޯ�ݸޔԍ�(VP�\��)
		rsd("LineCode")      = ""				'�q�H�R�[�h(VP�\��)
		rsd("HeadID")        = ""				'�w�b�h�h�c(ICCT�Ή�)03/07/31
		rsd("YLocExt")       = sYLocExt			'�O�����W I20040209
	End If

	rsd.Update
	rsd.Close
	sErrMsg = ""

End Function

'�\��`�F�b�N
Function AppUpdCheck(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
			sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, _
			sOpeRec, sOpeDel, sTerm_CD, sNinID, _
			iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, iCnt, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�\�񎞊ԑсiHH�j
'	sRDType			[ I ]��ށiDUAL,REC,DEL,�󔒁j
'	sContNoRec 		[ I ]�����R���e�i�ԍ�
'	sBKNo			[ I ]�����u�b�L���O�ԍ�
'	sContSizeRec	[ I ]�����R���e�i�T�C�Y
'	bChkA			[ I ]�����V���[�V�ɔ��o���ڂ��Ȃ�
'	bChkB			[ I ]�����V���[�V��[�ςɎg�p
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sContNoDel 		[ I ]���o�R���e�i�ԍ�
'	sChID			[ I ]���o�V���[�VID
'	sBLNo 			[ I ]���o�a�k�ԍ�
'	sVslCode		[ O ]���o�D���i�a�k�w�莞�j
'	sVoyage			[ O ]���o���q�i�a�k�w�莞�j
'	sContSizeDel	[I/O]���o�R���e�i�T�C�Y
'	sOpeRec			[ O ]�����I�y���[�^
'	sOpeDel			[ O ]���o�I�y���[�^
'	sTerm_CD		[ I ]�^�[�~�i��
'   sNinID			[ I ]�F�؂h�c
'	iAssignCnt 		[ I ]�O���[�v������
'	iOpenSlot 		[ I ]�J���g��
'	iDualOpeNo		[ O ]�f���A����Ɣԍ�
'	iOpeOrder		[ O ]��Ə���
'	iCnt			[ O ]�\��g��
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheck = False

	'���o���̃`�F�b�N(2003/08/27 ���Ұ��ǉ� sNinID ICCT�Ή�)
	If sRDType = "DEL" or sRDType = "DUAL" Then
		If Not AppUpdCheckDel(conn, rsd, _
						sGrpID, _
						sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, sOpeDel, _
						bChkC, sTerm_CD, sNinID, sErrMsg) Then
			Exit Function
		End If
	End If

	'�������̃`�F�b�N(02/02/09 ���Ұ��ǉ� sVsl, sVoy)
	If sRDType = "REC" or sRDType = "DUAL" Then
		If Not AppUpdCheckRec(conn, rsd, _
						sContNoRec, sBKNo, sContSizeRec, sOpeRec,  _
						sVslCode, sVoyage, sErrMsg) Then
			Exit Function
		End If
	End If

	'���o�����̃`�F�b�N
	If sRDType = "DUAL" Then
		'���p�V���[�V�łȂ��ꍇ�̓T�C�Y�̃}�b�`���`�F�b�N(���̃R�����g�̑Ή�������2001/6/2 M.OKIMI)
		If (Not bChkC) Then
			If ((sContSizeRec =  "20" and sContSizeDel <> "20") or _
				(sContSizeRec <> "20" and sContSizeDel =  "20")) Then
				sErrMsg = "���o���̃R���e�i�T�C�Y���قȂ�܂�"
				Exit Function
			End If
		End If
	End If

	'���ԑюw��̏ꍇ�͘g���`�F�b�N
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'�\��g�ɂ��`�F�b�N
		If Not AppUpdCheckTerm(conn, rsd, _
						sUsrID, sGrpID, _
						sYMD, sHH, _
						sRDType, _
						sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
						sContNoDel, sChID, sBLNo, sContSizeDel, _
						iAssignCnt, iOpenSlot, sTerm_CD, _
						iDualOpeNo, iOpeOrder, iCnt, sErrMsg) Then

			Exit Function
		End If
	Else
		iDualOpeNo = 0		'�f���A����Ɣԍ�
		iOpeOrder = 0		'��Ə���
		iCnt = 0			'�\��g��
	End If

	AppUpdCheck = True
End Function

'���o���\��`�F�b�N
Function AppUpdCheckDel(conn, rsd, _
			sGrpID, _
			sContNoDel, sChID, sBLNo, sVslCode, sVoyage, sContSizeDel, sOpeDel, _
			bChkC, sTerm_CD, sNinID, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sGrpID			[ I ]�O���[�vID
'	sContNoDel 		[ I ]���o�R���e�i�ԍ�
'	sChID			[ I ]���o�V���[�VID
'	sBLNo 			[ I ]���o�a�k�ԍ�
'	sVslCode		[ O ]���o�D���i�a�k�w�莞�j
'	sVoyage			[ O ]���o���q�i�a�k�w�莞�j
'	sContSizeDel	[I/O]���o�R���e�i�T�C�Y
'	sOpeDel			[ O ]���o�I�y���[�^
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sTerm_CD  		[ I ]�^�[�~�i��
'   sNinID			[ I ]�F��ID
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheckDel = False

	sVslCode = ""
	sVoyage  = ""

	If sContNoDel <> "" Then	'�R���e�i�w��
		'�\���̗L���`�F�b�N�i�w��R���e�i�j
		If ChkAppInfoExt(conn, rsd, sContNoDel) Then
			sErrMsg = "���ɗ\��ς݂̃R���e�i�ł��i" & sContNoDel & "�j"
			Exit Function
		End If

		'���o���R���e�i�̎擾
		Call GetOutCont(conn, rsd, sContNoDel)
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "���o�R���e�i���݌ɂ��Ă��܂���i" & sContNoDel & "�j"
			Exit Function
		End IF
'VP�Ή�
'''		If  Trim(rsd("Terminal")) <> "KA" Then
'''			rsd.Close
'''			sErrMsg = "���o�R���e�i�����łɍ݌ɂ��Ă��܂���i" & sContNoDel & "�j"
'''			Exit Function
'''		End If

		If  sTerm_CD = "KA" Then
			If Trim(rsd("Terminal")) <> "KA" Then
				rsd.Close
				sErrMsg = "���o�R���e�i�����łɍ݌ɂ��Ă��܂���i" & sContNoDel & "�j"
				Exit Function
			End If
		End If
'<2003/08/25 ICCT�Ή�>
'''		If  sTerm_CD = "HK" Then
'''			If  Trim(rsd("Terminal")) <> "HK" Then
'''				rsd.Close
'''				sErrMsg = "���o�R���e�i������ɍ݌ɂ��Ă��܂���i" & sContNoDel & "�j"
'''				Exit Function
'''			End If
'''		End If

   		If  sTerm_CD = "IC" Then
   			If  Trim(rsd("Terminal")) <> "IC" Then
   				rsd.Close
   				sErrMsg = "���o�R���e�i���h�b�b�s�ɍ݌ɂ��Ă��܂���i" & sContNoDel & "�j"
   				Exit Function
   			End If
		End If

		If  Trim(rsd("FullEmpty")) = "F" Then
			If IsNull(rsd("DelOKDate")) Then
				rsd.Close
				sErrMsg = "�w��R���e�i�͔��o�s�ł��i" & sContNoDel & "�j"
				Exit Function
			End If
			If sYMD < GetYMDStr(rsd("DelOKDate")) Then
				rsd.Close
				sErrMsg = "�w��R���e�i�͔��o�s�ł��i" & sContNoDel & "�j"
				Exit Function
			End If
			If not IsNull(rsd("DemFTDate")) Then
				If sYMD > GetYMDStr(rsd("DemFTDate")) Then
					rsd.Close
					sErrMsg = "�w��R���e�i�͔��o�s�ł��i" & sContNoDel & "�j"
					Exit Function
				End If
			End If
			If not IsNull(rsd("OLTFrom"))  Then	
				If GetYMDStr(rsd("OLTFrom")) > sYMD or _
				   GetYMDStr(rsd("OLTTo"))   < sYMD Then
					rsd.Close
					sErrMsg = "�w��R���e�i�͔��o�s�ł��i" & sContNoDel & "�j"
					Exit Function
				End If
			End If
		End If
'���[�t�@�[�R���e�i����o�\�ɂ���(01/09/04)
'''		If  Left( Trim(rsd("RHO")), 1) = "R" Or _
'''			Right(Trim(rsd("RHO")), 1) = "R" Then

'''			rsd.Close
'''			sErrMsg = "���[�t�@�[�R���e�i�̈וs�ł��i" & sContNoDel & "�j"
'''			Exit Function
'''		End If

'�F�؂h�c�ɒl������ꍇ�͓��͂����l�Ɠ������Ȃ��ƃG���[(2003/08/27)

		If  Trim(rsd("DONo")) <> "" Then
			If  Trim(rsd("DONo")) <> sNinID  Then
   				rsd.Close
   				sErrMsg = "�F�؂h�c�𐳂������͂��Ă��������i" & sContNoDel & "�j"
   				Exit Function
			End If
   		End If

		sContSizeDel = Trim(rsd("ContSize"))
		sOpeDel = Trim(rsd("Operator"))
'02/02/09 ADD
		sVslCode = Trim(rsd("VslCode"))
		sVoyage  = Trim(rsd("Voyage"))
'02/02/09 ADD
		rsd.Close
	Else						'�a�k�w��
		'���o���}�j�t�F�X�g�R���e�i�̎擾(�����b�x�Ή�(01/10/01)
		Call GetOutBLContCont(conn, rsd, sBLNo, sTerm_CD)
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "�a�k�ԍ�������������܂���i" & sBLNo & "�j"
			Exit Function
		End IF
		sVslCode = Trim(rsd("VslCode"))
		sVoyage  = Trim(rsd("Voyage"))
		sOpeDel = Trim(rsd("Operator"))

		'�d���`�F�b�N
		Do Until rsd.EOF
			If sVslCode <> Trim(rsd("VslCode")) or _
			   sVoyage  <> Trim(rsd("Voyage")) Then
				rsd.Close
				sErrMsg = "����̂a�k�ԍ������݂��܂��̂ŃR���e�i�ԍ��w��ŗ\�񂵂ĉ�����"
				Exit Function
			End If
			rsd.MoveNext
		Loop

		rsd.Close
	End If

	'�V���[�V���`�F�b�N
	If sChID <> "" Then	'���o�V���[�VID
		'�V���[�V�̎擾
	 	Call GetChassis(conn, rsd, sChID)
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "�V���[�V���o�^����Ă��܂���i" & sChID & "�j"
			Exit Function
		End IF
		If Trim(rsd("GroupID")) <> sGrpID Then
			rsd.Close
			sErrMsg = "�V���[�V�̏������قȂ�܂��i" & sChID & "�j"
			Exit Function
		End If
		If rsd("MixSizeFlag") <> "Y" Then	'���p�V���[�V�łȂ�
			If bChkC Then
				rsd.Close
				sErrMsg = "20/40���p�V���[�V�ł͂���܂���i" & sChID & "�j"
				Exit Function
			End If
			If (sContSizeDel <> "20" and rsd("Size20Flag") =  "Y") or _
			   (sContSizeDel =  "20" and rsd("Size20Flag") <> "Y") Then
				rsd.Close
				sErrMsg = "�V���[�V�T�C�Y���قȂ�܂��i" & sChID & "�j"
				Exit Function
			End If
		End If
		rsd.Close

		'�w��V���[�V�������Ɠ��ɗ\�񂳂�Ă��邩�`�F�b�N
		Call GetAppInfoChassis(conn, rsd, sYMD, sChID)
		If Not rsd.EOF Then
			rsd.Close
			sErrMsg = "�w��V���[�V�͂��łɗ\�񂳂�Ă��܂��i" & sChID & "�j"
			Exit Function
		End If
		rsd.Close
	End If

	AppUpdCheckDel = True
End Function

'�������\��`�F�b�N(02/02/09 �p�����[�^ADD)
Function AppUpdCheckRec(conn, rsd, _
			sContNoRec, sBKNo, sContSizeRec, sOpeRec, _
			sVslCode, sVoyage, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sContNoRec 		[ I ]�����R���e�i�ԍ�
'	sBKNo			[ I ]�����u�b�L���O�ԍ�
'	sContSizeRec	[ I ]�����R���e�i�T�C�Y
'	sOpeRec			[ O ]�����I�y���[�^
'	sVslCode		[ O ]�����D��		'02/02/09
'	sVoyage			[ O ]�������q		'02/02/09
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheckRec = False

	'�u�b�L���O�̗L������(02/02/09 ADD)
	If Not ChkBooking(conn, rsd, sBKNo, sOpeRec, sVslCode, sVoyage) Then
		sErrMsg = "�u�b�L���O�����݂��܂���i" & sBKNo & "�j"
		Exit Function
	End If

	'�\���̗L���`�F�b�N�i�w��R���e�i�j
	If ChkAppInfoExt(conn, rsd, sContNoRec) Then
		sErrMsg = "���ɗ\��ς݂̃R���e�i�ł��i" & sContNoRec & "�j"
		Exit Function
	End If

	AppUpdCheckRec = True

End Function

'�\��g�ɂ��`�F�b�N
Function AppUpdCheckTerm(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
			sContNoDel, sChID, sBLNo, sContSizeDel, _
			iAssignCnt, iOpenSlot, sTerm_CD, iDualOpeNo, iOpeOrder, iCnt, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�\�񎞊ԑсiHH�j
'	sRDType			[ I ]��ށiDUAL,REC,DEL,�󔒁j
'	sContNoRec 		[ I ]�����R���e�i�ԍ�
'	sBKNo			[ I ]�����u�b�L���O�ԍ�
'	sContSizeRec	[ I ]�����R���e�i�T�C�Y
'	bChkA			[ I ]�����V���[�V�ɔ��o���ڂ��Ȃ�
'	bChkB			[ I ]�����V���[�V��[�ςɎg�p
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sContNoDel 		[ I ]���o�R���e�i�ԍ�
'	sChID			[ I ]���o�V���[�VID
'	sBLNo 			[ I ]���o�a�k�ԍ�
'	sContSizeDel	[I/O]���o�R���e�i�T�C�Y
'	iAssignCnt 		[ I ]�O���[�v������
'	iOpenSlot 		[ I ]�J���g��
'	sTerm_CD		[ I ]���o����^�[�~�i���R�[�h
'	iDualOpeNo		[ O ]�f���A����Ɣԍ�
'	iOpeOrder		[ O ]��Ə���
'	iCnt			[ O ]���݉����񐔁i���o�^�O���[�v�g���j
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheckTerm = False

	'�\�����̎擾�i�w����A�w�莞�ԑсA�w��O���[�v�j
	Call GetAppInfoTerm(conn, rsd, sYMD, sHH, sGrpID)
	iCnt = 0			'���݉����񐔁i���o�^�O���[�v�g���j
	iDualOpeNo = 0		'�f���A����Ɣԍ�
	iOpeOrder = 0		'��Ə���
	Do Until rsd.EOF
		'�����Ńf���A���ɂȂ��ꍇ(���͂��ꂽ�̂������ŁE�Y���̔��o�����{��)
		If sRDType = "REC" and (Not bChkA) and (Not bChkB) Then
			'�y�A�ɂȂ蓾����o���`�F�b�N
				'01/10/19 sta
				'''			If rsd("DualOpeNo") = 0 and _
				'VP�Ή�        		   rsd("RecDel") = "D" and _
				'''			   Trim(rsd("ChassisId")) = "" Then
				'01/10/19 sta
'02/04/27 sta(DUAL�͓���^�[�~�i���ɂ���)
	   			'If rsd("DualOpeNo") = 0 and rsd("RecDel") <> "R" and _
	   			'   Trim(rsd("ChassisId")) = "" Then
				'	'���������ŁE���肩�܂��́A����ȊO�̑O�󂯂̏ꍇ�͓���VP�^�[�~�i���̏ꍇ
				'	If (sTerm_CD = "KA" OR sTerm_CD = "HK") or _
				'	   (sTerm_CD <> "KA" AND sTerm_CD <> "HK" AND sTerm_CD = Trim(rsd("Terminal")))	  Then	
				'		'���p���T�C�Y���}�b�`����ꍇ
				'		If bChkC or _
				'		   (rsd("ContSize") =  "20" and sContSizeRec =  "20") or _
				'		   (rsd("ContSize") <> "20" and sContSizeRec <> "20") Then
				'			iDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
				'			iOpeOrder = rsd("OpeOrder") - 1
				'		End If
				'	End If
				'End If

   			If rsd("DualOpeNo") = 0 and rsd("RecDel") <> "R" and _
   			   Trim(rsd("ChassisId")) = "" Then
				'���������ŁE���肩�܂��́A����ȊO�̑O�󂯂̏ꍇ�͓���VP�^�[�~�i���̏ꍇ
				'If (sTerm_CD = "KA" OR sTerm_CD = "HK") or _
				'   (sTerm_CD <> "KA" AND sTerm_CD <> "HK" AND sTerm_CD = Trim(rsd("Terminal")))	  Then	
				'02/04/27 ����^�[�~�i����DUAL�ɂ���
				If sTerm_CD = Trim(rsd("Terminal"))	  Then	
					'���p���T�C�Y���}�b�`����ꍇ
					If bChkC or _
					   (rsd("ContSize") =  "20" and sContSizeRec =  "20") or _
					   (rsd("ContSize") <> "20" and sContSizeRec <> "20") Then
						iDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
						iOpeOrder = rsd("OpeOrder") - 1
					End If
				End If
			End If
		End If
		'���o�Ńf���A���ɂȂ��ꍇ(���͂��ꂽ�̂����o�ŁE�Y���̔��������{��)
		If sRDType = "DEL" and sChID = "" Then	
			'�y�A�ɂȂ蓾��������`�F�b�N
			If rsd("DualOpeNo") = 0 and _
			   rsd("RecDel") = "R" and _
			   rsd("NotDelFlag") <> "Y" and _
			   rsd("NightFlag") <> "Y" Then
				'���ł�����̂ݑΏۂƂȂ�
'2003/08/25
'''''			If Trim(rsd("Terminal")) = "KA" or Trim(rsd("Terminal")) = "HK" Then
				If Trim(rsd("Terminal")) = "KA" or Trim(rsd("Terminal")) = "IC" Then
'2003/08/25
				'02/04/27 ����^�[�~�i����DUAL�ɂ���
					If sTerm_CD = Trim(rsd("Terminal"))	  Then	
						'���p���T�C�Y���}�b�`����ꍇ
						If rsd("MixSizeFlag") = "Y" or _
						   (rsd("ContSize") =  "20" and sContSizeDel =  "20") or _
						   (rsd("ContSize") <> "20" and sContSizeDel <> "20") Then
							iDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
							iOpeOrder = rsd("OpeOrder") + 1
						End If
					End If
				End If
			End If
		End If
		If rsd("DualOpeNo") = 0 or _
		   rsd("RecDel") <> "R" Then				'=D---> <>R�����ȊO�ɕύX(VP�Ή�)
			'���o�͒P�Ƃ̂݃J�E���g�������񐔂̃J�E���g
			iCnt = iCnt + 1	'���݉����񐔁i���o�^�O���[�v�g���j
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	If iDualOpeNo = 0 and _
	   iOpenSlot = 0 and _
	   iCnt >= iAssignCnt Then
		sErrMsg = "���蓖�Ęg������܂���"
		Exit Function
	End If

	'���̍�Ƃƃf���A���ɂȂ�Ȃ��ꍇ�͍�Ə��ʂ��v�Z
	If iOpeOrder = 0 Then
		'�V�K��Ə��ʂ̎擾�i�w����A�w�莞�ԑсj
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, sRDType)
	End If

	AppUpdCheckTerm = True

End Function
%>