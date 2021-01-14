<% 

'�\��ړ�����
Function UpdOpeMov(conn, rsd, sOpeNo, sYMD, sHH, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sOpeNo			[ I ]��Ɣԍ�
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�ړ����ԑсiHH�j
'	sErrMsg			[ O ]�G���[���b�Z�[�W

	Dim sStatus, sDelFlag, sLockFlag, iDualOpeNo, sGrpID, sUsrID
	Dim iNewDualOpeNo, iOpeOrder, sTerm_CD
	Dim sRDType, sContSize, bChkA, bChkB, bChkC, sChID, sPlace
'2001/03/26�I�����E��̍�Ə���DUAL
	Dim iOpeOrder_M, iOpeOrder_S, iDualOpeNo_S, sDual

	sErrMsg = "�c�a�X�V�G���["

	'�\�����̎擾�i�w���Ɣԍ��j
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "�Ώۂ��폜����Ă��܂�"
		Exit Function
	End If
	sStatus    = Trim(rsd("Status"))			'���
	sDelFlag   = Trim(rsd("DelFlag"))			'�폜�t���O
	sLockFlag  = Trim(rsd("LockFlag"))			'���b�N�t���O
	iDualOpeNo =      rsd("DualOpeNo")			'�f���A����Ɣԍ�
	sGrpID     = Trim(rsd("GroupID"))			'�O���[�vID
	sUsrID     = Trim(rsd("UserID"))			'���[�UID
	sRDType    = Trim(rsd("RecDel"))			'���o���敪
	sContSize  = Trim(rsd("ContSize"))			'�R���e�i�T�C�Y
	bChkA      = Trim(rsd("NotDelFlag"))  = "Y"	'���o���ڂ��Ȃ�
	bChkB      = Trim(rsd("NightFlag"))   = "Y"	'�[�σV���[�V
	bChkC      = Trim(rsd("MixSizeFlag")) = "Y"	'20/40���p�V���[�V
	sChID      = Trim(rsd("ChassisId"))			'����ID
	sPlace     = Trim(rsd("Place"))				'�ꏊ
	iOpeOrder_M =     rsd("OpeOrder")			'��Ə���(�I�����ꂽ��)(3/26)
	sTerm_CD   = Trim(rsd("Terminal"))			'�^�[�~�i��
	rsd.Close

	'���ݏ�Ԃ̃`�F�b�N
	If sStatus = "04" Then	'�����̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��������Ă��܂�"
		Exit Function
	End If

	If sStatus = "02" Then	'�\�񒆂̏ꍇ
		If sLockFlag = "Y" Then	'���b�N
			sErrMsg = "�Ώۂ͍�Ƃ��m�肵�Ă��܂�"
			Exit Function
		End If
	End If

	'�����ŃV���[�V�����肵�Ă���i���r�x�����ς݁j�̏ꍇ
	If sRDType = "R" And sChID <> "" Then
		'�V���[�V�̎擾
		Call GetChassis(conn, rsd, sChID)
		If Trim(rsd("NotDelFlag")) = "Y" Then		'���o���ڂ��Ȃ�
			bChkA = True
		End If
		If Trim(rsd("NightFlag")) = "Y" Then		'�[�σV���[�V
			bChkB = True
		End If
		bChkC = Trim(rsd("MixSizeFlag")) = "Y"		'20/40���p�V���[�V
		rsd.Close
	End If

'2001/03/26	START	DUAL�̏ꍇ���葤�ɍ�Ɣԍ������č�Ə��ʂ������ꍇ��DUAL�I��
	sDual = ""
	If iDualOpeNo > 0 Then 
		'�\�����̎擾�i�w���Ɣԍ��j
		Call GetAppInfoOpeNo(conn, rsd, CLng(iDualOpeNo))
		If rsd.EOF Then
			rsd.Close
			sErrMsg = "�Ώۂ��폜����Ă��܂�"
			Exit Function
		End If
		iDualOpeNo_S =     rsd("DualOpeNo")		'�f���A����Ɣԍ�(�I�����ꂽ����)
		iOpeOrder_S  =     rsd("OpeOrder")		'��Ə���(�I�����ꂽ����)
		rsd.Close

'DUALOPE�ԍ������݂ɓ����ō�Ə��ʂ������ꍇDUAL�I��
		If iOpeOrder_M = iOpeOrder_S And CLng(sOpeNo) = iDualOpeNo_S Then		
			sDual = "1"
		End If
	
	End iF
'2001/03/26	End

	'�ړ��v�Z
	Call UpdOpeMovCalc(conn, rsd, sOpeNo, sGrpID, sUsrID, sYMD, sHH, _
						sRDType, _
						sContSize, bChkA, bChkB, bChkC,  _
						sChID, sPlace, sDual, sTerm_CD, _
						iNewDualOpeNo, iOpeOrder, sErrMsg)
	If sErrMsg <> "" Then
		Exit Function
	End If

	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	rsd("UpdtTime")  = Now()				'�X�V����
	rsd("UpdtPgCd")  = "SYWB031"			'�X�V��۸��і�
	rsd("UpdtTmnl")  = "WEB"				'�X�V�[����
	rsd("Status")    = "02"					'���
	rsd("DelFlag")   = ""					'�폜�׸
	rsd("CReason")   = ""					'�L�����Z�����R
	rsd("Term")      = sHH					'���ԑ�
'2001/03/27DUAL���ʂ������Ă��ċ敪�������̎��͍�Ə��ʂ��v���X�P����
	If sDual =  "1" And  sRDType = "R" And _
		sHH <> "12" and sHH <> "A" and sHH <> "B" Then
			iOpeOrder = Cint(iOpeOrder) + 1		'��Ə��ʌv�Z
	End If
	rsd("OpeOrder")  = iOpeOrder			'��Ə���

	rsd("HopeTerm")  = sHH					'��]���ԑ�
'2001/03/27DUAL���ʂ������Ȃ����̂̂ݍX�V	
	If sDual = "" Then						
		rsd("DualOpeNo") = iNewDualOpeNo	'�f���A����Ɣԍ�
	End If
'2001/03/27DUAL���ʂ������Ȃ����̂̂ݍX�V	
	rsd("LockFlag")  = ""					'���b�N�t���O
	rsd("SendFlag")  = "Y"					'���M�t���O
	rsd.Update
	rsd.Close

'2001/03/27DUAL���ʂ������Ȃ��ꍇ�͈ȉ��̏���
	If sDual = "" Then						
		If iDualOpeNo > 0 Then	'�f���A����Ɣԍ�
			Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
			rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
			rsd("UpdtTime")  = Now()			'�X�V����
			rsd("UpdtPgCd")  = "SYWB031"		'�X�V��۸��і�
			rsd("UpdtTmnl")  = "WEB"			'�X�V�[����
			rsd("SendFlag")  = "Y"				'���M�t���O
			rsd.Update
			rsd.Close
		End If
		If iNewDualOpeNo > 0 Then	'�f���A����Ɣԍ�
			Call GetAppInfoOpeNoUpd(conn, rsd, iNewDualOpeNo)
			rsd("DualOpeNo") = CLng(sOpeNo)		'�f���A����Ɣԍ�
			rsd("UpdtTime")  = Now()			'�X�V����
			rsd("UpdtPgCd")  = "SYWB031"		'�X�V��۸��і�
			rsd("UpdtTmnl")  = "WEB"			'�X�V�[����
			rsd("SendFlag")  = "Y"				'���M�t���O
			rsd.Update
			rsd.Close
		End If
	Else
'2001/03/27DUAL���ʂ����ꍇ�͈ȉ��̏���
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
		rsd("UpdtTime")  = Now()			'�X�V����
		rsd("UpdtPgCd")  = "SYWB031"		'�X�V��۸��і�
		rsd("UpdtTmnl")  = "WEB"			'�X�V�[����
		rsd("Status")    = "02"				'���
		rsd("DelFlag")   = ""				'�폜�׸
		rsd("CReason")   = ""				'�L�����Z�����R
		rsd("Term")      = sHH				'���ԑ�
		rsd("HopeTerm")  = sHH				'��]���ԑ�
		rsd("LockFlag")  = ""				'���b�N�t���O
		rsd("SendFlag")  = "Y"				'���M�t���O
		rsd("OpeOrder")  = iOpeOrder		'��Ə���
		rsd.Update
		rsd.Close
	End If

'test
'If sDual <> "" Then						
'sErrMsg = "�e�X�g�� sDual= " & sDual & "��Ə���=" & iOpeOrder
'Exit Function
'End If
'test

	sErrMsg = ""

End Function

'�ړ��v�Z
Function UpdOpeMovCalc(conn, rsd, sOpeNo, sGrpID, sUsrID, sYMD, sHH, _
						sRDType, _
						sContSize, bChkA, bChkB, bChkC,  _
						sChID, sPlace, sDual, sTerm_CD, _
						iNewDualOpeNo, iOpeOrder, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sOpeNo			[ I ]��Ɣԍ�
'	sGrpID     		[ I ]�O���[�vID
'	sUsrID     		[ I ]���[�UID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�ړ����ԑсiHH�j
'	sRDType			[ I ]��ށiD,R,V�j
'	sContSize		[ I ]�R���e�i�T�C�Y
'	bChkA			[ I ]�����V���[�V�ɔ��o���ڂ��Ȃ�
'	bChkB			[ I ]�����V���[�V��[�ςɎg�p
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sChID			[ I ]���o�V���[�VID�^�����V���[�VID
'	sPlace			[ I ]�ꏊ
'	sDual			[ I ]�c�t�`�k����					3/27 add
'	sTerm_CD		[ I }�Ώۃ^�[�~�i��					10/4 add
'	iNewDualOpeNo	[ O ]�f���A����Ɣԍ�
'	iOpeOrder		[ O ]��Ə���
'	sErrMsg			[ O ]�G���[���b�Z�[�W

	Dim sWk, iAssignCnt, bOk, iCnt
	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime

	sErrMsg = ""

	'*** �\����A�\�񎞊ԑт��L�����`�F�b�N ***
	'���ݓ��t�擾
	sDateNow = GetYMDStr(Date())

	If sDateNow > sYMD Then
		sErrMsg = "��Ɠ����K�؂ł͂���܂���"
		Exit Function
	End If
	If sDateNow = sYMD Then	'���ݓ��w��̏ꍇ
		'�^�s�󋵂��擾
		Call GetOpeStatus(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime)
		sWk = TermNumToSym(iNextTime)
		If sHH < sWk or (sHH = sWk and iNextStat = 2) Then
			sErrMsg = "�w�莞�ԑт̗\��͒��ߐ؂��Ă��܂�"
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
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'�O���[�v�̊��蓖�Ęg���擾
		iAssignCnt = GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)
	End If

	'���ԑюw��̏ꍇ�͘g���`�F�b�N
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'�\��g�ɂ��`�F�b�N
		If Not AppUpdCheckTerm(conn, rsd, _
						sUsrID, sGrpID, _
						sYMD, sHH, _
						sRDType, _
						sContSize, bChkA, bChkB, bChkC,  _
						sChID, sPlace, sDual, sTerm_CD, _
						iAssignCnt, iOpenSlot, iNewDualOpeNo, iOpeOrder, sErrMsg) Then
			Exit Function
		End If
	Else
		iNewDualOpeNo = 0		'�f���A����Ɣԍ�
		iOpeOrder = 0			'��Ə���
	End If

End Function

'�\��g�ɂ��`�F�b�N
Function AppUpdCheckTerm(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContSize, bChkA, bChkB, bChkC,  _
			sChID, sPlace, sDual, sTerm_CD, _
			iAssignCnt, iOpenSlot, iNewDualOpeNo, iOpeOrder, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�\�񎞊ԑсiHH�j
'	sRDType			[ I ]��ށiD,R,V�j
'	sContSize		[ I ]�R���e�i�T�C�Y
'	bChkA			[ I ]�����V���[�V�ɔ��o���ڂ��Ȃ�
'	bChkB			[ I ]�����V���[�V��[�ςɎg�p
'	bChkC  			[ I ]�����V���[�V��20/40���p�V���[�V
'	sChID			[ I ]���o�V���[�VID�^�����V���[�VID
'	sPlace			[ I ]�ꏊ
'	sDual			[ I ]DUAL����		'2001/03/27
'	sTerm_CD		[ I ]�Ώۃ^�[�~�i�� '2001/10/04
'	iAssignCnt 		[ I ]�O���[�v������
'	iOpenSlot 		[ I ]�J���g��
'	iNewDualOpeNo	[ O ]�f���A����Ɣԍ�
'	iOpeOrder		[ O ]��Ə���
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	Dim iCnt	'���݉����񐔁i���o�^�O���[�v�g���j

	AppUpdCheckTerm = False

	'�\�����̎擾�i�w����A�w�莞�ԑсA�w��O���[�v�j
	Call GetAppInfoTerm(conn, rsd, sYMD, sHH, sGrpID)
	iCnt = 0	'���݉����񐔁i���o�^�O���[�v�g���j
	iNewDualOpeNo = 0		'�f���A����Ɣԍ�
	iOpeOrder = 0		'��Ə���
	Do Until rsd.EOF
'2001/03/27 DUAL���ʂ������ĂȂ��ꍇ�݈̂ȉ��̏���������
		If sDual = "" Then
			'�����Ńf���A���ɂȂ��ꍇ(�I�����ꂽ�̂������ŁE�Y���̔��o�����{��)
			If sRDType = "R" and (Not bChkA) and (Not bChkB) Then
									'�y�A�ɂȂ蓾����o���`�F�b�N
					'�Y�����鑊�肪�A���o���܂��͋�o���̎��ɕύX����
					'''				If rsd("DualOpeNo") = 0 and _
					'''VP�Ή�		   rsd("RecDel") = "D" and _
					'''				   (Trim(rsd("ChassisId")) = "" or _
					'''				    Trim(rsd("ChassisId")) = sChID) Then
					'If rsd("DualOpeNo") = 0 and _
					'   rsd("RecDel") <> "R" and _			
					'   (Trim(rsd("ChassisId")) = "" or _
					'    Trim(rsd("ChassisId")) = sChID) Then
					'	'���������ŁE���肩�܂��́A����ȊO�̑O�󂯂̏ꍇ�͓���VP�^�[�~�i���̏ꍇ
					'	If (sTerm_CD = "KA" OR sTerm_CD = "HK") or _
					'	   (sTerm_CD <> "KA" AND sTerm_CD <> "HK" AND sTerm_CD = Trim(rsd("Terminal")))	  Then	
					'		'���p���T�C�Y���}�b�`����ꍇ
					'		If bChkC or _
					'		   (rsd("ContSize") =  "20" and sContSize =  "20") or _
					'		   (rsd("ContSize") <> "20" and sContSize <> "20") Then
					'			iNewDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
					'			iOpeOrder = rsd("OpeOrder") - 1
					'		End If
					'	End If
					'End If
				'02/04/27 (DUAL�͓���^�[�~�i���ɂ���) sta
				If rsd("DualOpeNo") = 0 and _
				   rsd("RecDel") <> "R" and _			
				   (Trim(rsd("ChassisId")) = "" or _
				    Trim(rsd("ChassisId")) = sChID) Then
					'����VP�^�[�~�i��
					If sTerm_CD = Trim(rsd("Terminal"))	  Then	
						'���p���T�C�Y���}�b�`����ꍇ
						If bChkC or _
						   (rsd("ContSize") =  "20" and sContSize =  "20") or _
						   (rsd("ContSize") <> "20" and sContSize <> "20") Then
							iNewDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
							iOpeOrder = rsd("OpeOrder") - 1
						End If
					End If
				End If
				'02/04/27 (DUAL�͓���^�[�~�i���ɂ���) end
			End If

			'���o�Ńf���A���ɂȂ��ꍇ(�I�����ꂽ�̂����o�ŁE�Y���̔��������{��)
			'�����ȊO�Ńf���A���ɂȂ��ꍇ(�I�����ꂽ�̂������ȊO�ŁE�Y���̔��������{��)
'''			If sRDType = "D" Then	
			If sRDType <> "R" Then			'�����łȂ�
				'�y�A�ɂȂ蓾��������`�F�b�N
				If rsd("DualOpeNo") = 0 and _
				   rsd("RecDel") = "R" and _
				   rsd("NotDelFlag") <> "Y" and _
				   rsd("NightFlag") <> "Y" and _
				   (sChID = "" or _
				    Trim(rsd("ChassisId")) = sChID) Then
						'�����^�[�~�i���Ɍ���
						If sTerm_CD = Trim(rsd("Terminal")) Then	
							'���p���T�C�Y���}�b�`����ꍇ
							If rsd("MixSizeFlag") = "Y" or _
							   (rsd("ContSize") =  "20" and sContSize =  "20") or _
							   (rsd("ContSize") <> "20" and sContSize <> "20") Then
								iNewDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
								iOpeOrder = rsd("OpeOrder") + 1
							End If
						End If
				End If
			End If
		End If

'''		If rsd("DualOpeNo") = 0 or _
'''		   rsd("RecDel") = "D" Then
'''		   rsd("RecDel") <> "R" Then	'02/04/27 sywb012.inc�ɍ��킹��
		If rsd("DualOpeNo") = 0 or rsd("RecDel") <> "R" Then
			'���o�͒P�Ƃ̂݃J�E���g�������񐔂̃J�E���g
			iCnt = iCnt + 1	'���݉����񐔁i���o�^�O���[�v�g���j
		End If

		rsd.MoveNext
	Loop
	rsd.Close

	If iNewDualOpeNo = 0 and _
	   iOpenSlot = 0 and _
	   iCnt >= iAssignCnt Then
		sErrMsg = "���蓖�Ęg������܂���"
		Exit Function
	End If

	'���̍�Ƃƃf���A���ɂȂ�Ȃ��ꍇ�͍�Ə��ʂ��v�Z
	'�܂���DUAL���ʃt���O�������Ă���ꍇ�͏��ʂ��v�Z����
	If iOpeOrder = 0 Or sDual <> "" Then
		'�V�K��Ə��ʂ̎擾�i�w����A�w�莞�ԑсj
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, sRDType)
	End If

	AppUpdCheckTerm = True

End Function

%>