<% 

'�\���������
Function UpdOpeChg(conn, rsd, sOpeNo, sYMD, sChgOpeNo, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sOpeNo			[ I ]��Ɣԍ�
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sChgOpeNo		[ I ]�����Ώۍ�Ɣԍ�
'	sErrMsg			[ O ]�G���[���b�Z�[�W

	Dim sGrpID(1), sUsrID(1), dWorkDate(1), sHH(1)
	Dim iOpeOrder(1), sStatus(1), sDelFlag(1), sLockFlag(1), iDualOpeNo(1)
	Dim sRDType(1), sContSize(1), bChkA(1), bChkB(1), bChkC(1), sChID(1), sPlace(1)
	Dim AppDate(1), sTerminal(1)

	sErrMsg = "�c�a�X�V�G���["

	'�\�����̎擾�i�w���Ɣԍ��j
	Call GetAppInfoOpeNo(conn, rsd, CLng(sChgOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "�������肪���݂��܂���"
		Exit Function
	End If
	dWorkDate(0)  =      rsd("WorkDate")	'��Ɠ�
	sHH(0)    	  = Trim(rsd("Term"))		'���ԑ�
	iOpeOrder(0)  =      rsd("OpeOrder")	'��Ə���
	sStatus(0)    = Trim(rsd("Status"))		'���
	sDelFlag(0)   = Trim(rsd("DelFlag"))	'�폜�t���O
	sLockFlag(0)  = Trim(rsd("LockFlag"))	'���b�N�t���O
	iDualOpeNo(0) =      rsd("DualOpeNo")	'�f���A����Ɣԍ�
	sGrpID(0)     = Trim(rsd("GroupID"))	'�O���[�vID
	sUsrID(0)     = Trim(rsd("UserID"))		'���[�UID
	sRDType(0)    = Trim(rsd("RecDel"))		'���o���敪
	sContSize(0)  = Trim(rsd("ContSize"))	'�R���e�i�T�C�Y
	bChkA(0)      = Trim(rsd("NotDelFlag"))	'���o���ڂ��Ȃ�
	bChkB(0)      = Trim(rsd("NightFlag"))	'�[�σV���[�V
	bChkC(0)      = Trim(rsd("MixSizeFlag"))'20/40���p�V���[�V
	sChID(0)      = Trim(rsd("ChassisId"))	'����ID
	sPlace(0)     = Trim(rsd("Place"))		'�ꏊ
	AppDate(0)	  =      rsd("AppDate")		'�\����
	sTerminal(0)  = Trim(rsd("Terminal"))	'�^�[�~�i��
	rsd.Close

	If sDelFlag(0) = "Y" Then	'�폜�̏ꍇ
		sErrMsg = "�������肪�폜����Ă��܂�"
		Exit Function
	End If

	If sStatus(0) = "03" Then	'�L�����Z���̏ꍇ
		sErrMsg = "��������͍�Ƃ��L�����Z������Ă��܂�"
		Exit Function
	End If
	If sStatus(0) = "04" Then	'�����̏ꍇ
		sErrMsg = "��������͍�Ƃ��������Ă��܂�"
		Exit Function
	End If

	If sStatus(0) = "02" Then	'�\�񒆂̏ꍇ
		If sLockFlag(0) = "Y" Then	'���b�N
			sErrMsg = "��������͍�Ƃ��m�肵�Ă��܂�"
			Exit Function
		End If
	End If

	'�\�����̎擾�i�w���Ɣԍ��j
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "�Ώۂ��폜����Ă��܂�"
		Exit Function
	End If
	dWorkDate(1)  =      rsd("WorkDate")	'��Ɠ�
	sHH(1)    	  = Trim(rsd("Term"))		'���ԑ�
	iOpeOrder(1)  =      rsd("OpeOrder")	'��Ə���
	sStatus(1)    = Trim(rsd("Status"))		'���
	sDelFlag(1)   = Trim(rsd("DelFlag"))	'�폜�t���O
	sLockFlag(1)  = Trim(rsd("LockFlag"))	'���b�N�t���O
	iDualOpeNo(1) =      rsd("DualOpeNo")	'�f���A����Ɣԍ�
	sGrpID(1)     = Trim(rsd("GroupID"))	'�O���[�vID
	sUsrID(1)     = Trim(rsd("UserID"))		'���[�UID
	sRDType(1)    = Trim(rsd("RecDel"))		'���o���敪
	sContSize(1)  = Trim(rsd("ContSize"))	'�R���e�i�T�C�Y
	bChkA(1)      = Trim(rsd("NotDelFlag"))	'���o���ڂ��Ȃ�
	bChkB(1)      = Trim(rsd("NightFlag"))	'�[�σV���[�V
	bChkC(1)      = Trim(rsd("MixSizeFlag"))'20/40���p�V���[�V
	sChID(1)      = Trim(rsd("ChassisId"))	'����ID
	sPlace(1)     = Trim(rsd("Place"))		'�ꏊ
	AppDate(1)	  =      rsd("AppDate")		'�\����
	sTerminal(1)  = Trim(rsd("Terminal"))	'�^�[�~�i��
	rsd.Close

	If sDelFlag(1) = "Y" Then	'�폜�̏ꍇ
		sErrMsg = "�Ώۂ��폜����Ă��܂�"
		Exit Function
	End If

	If sStatus(1) = "03" Then	'�L�����Z���̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��L�����Z������Ă��܂�"
		Exit Function
	End If
	If sStatus(1) = "04" Then	'�����̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��������Ă��܂�"
		Exit Function
	End If

	If sStatus(1) = "02" Then	'�\�񒆂̏ꍇ
		If sLockFlag(1) = "Y" Then	'���b�N
			sErrMsg = "�Ώۂ͍�Ƃ��m�肵�Ă��܂�"
			Exit Function
		End If
	End If

	If sGrpID(0) <> sGrpID(1) Then
		sErrMsg = "��������̃O���[�v���قȂ�܂�"
		Exit Function
	End If
'VP�Ή��@����<->�����܂���(���o<->���o,���o<->��o��,��o��<->���o,��o��<->��o��)
'''	If sRDType(0) <> sRDType(1) Then
'''		sErrMsg = "��������̔��o���敪���قȂ�܂�"
'''		Exit Function
'''	End If
'�����̃`�F�b�N
	If (sRDType(0) = "R" and sRDType(1) <> "R") OR _
       (sRDType(0) <> "R" and sRDType(1) = "R")  Then
		sErrMsg = "��������̔��o���敪���قȂ�܂�"
		Exit Function
	End If

	If sContSize(0) <> sContSize(1) Then
		sErrMsg = "��������̃R���e�i�T�C�Y���قȂ�܂�"
		Exit Function
	End If
	If bChkA(0) <> bChkA(1) or _
	   bChkB(0) <> bChkB(1) or _
	   bChkC(0) <> bChkC(1) Then
		sErrMsg = "��������̃V���[�V�������قȂ�܂�"
		Exit Function
	End If

'�^�[�~�i���̃`�F�b�N
	If sTerminal(0) <> sTerminal(1) Then
		sErrMsg = "��������̃^�[�~�i�����قȂ�܂�"
		Exit Function
	End If

	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	rsd("UpdtTime")  = Now()				'�X�V����
	rsd("UpdtPgCd")  = "SYWB031"			'�X�V��۸��і�
	rsd("UpdtTmnl")  = "WEB"				'�X�V�[����
	rsd("WorkDate")  = dWorkDate(0)			'��Ɠ�
	rsd("Term")      = sHH(0)				'���ԑ�
	rsd("OpeOrder")  = iOpeOrder(0)			'��Ə���
	rsd("HopeTerm")  = sHH(0)				'��]���ԑ�
	rsd("DualOpeNo") = iDualOpeNo(0)		'�f���A����Ɣԍ�
	rsd("SendFlag")  = "Y"					'���M�t���O
'	rsd("AppDate") 	 = AppDate(0)			'�\����
	rsd.Update
	rsd.Close
	If iDualOpeNo(1) > 0 Then	'�f���A����Ɣԍ�
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo(1))
		rsd("DualOpeNo") = CLng(sChgOpeNo)	'�f���A����Ɣԍ�
		rsd("UpdtTime")  = Now()			'�X�V����
		rsd("UpdtPgCd")  = "SYWB031"		'�X�V��۸��і�
		rsd("UpdtTmnl")  = "WEB"			'�X�V�[����
		rsd("SendFlag")  = "Y"				'���M�t���O
		rsd.Update
		rsd.Close
	End If
	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sChgOpeNo))
	rsd("UpdtTime")  = Now()				'�X�V����
	rsd("UpdtPgCd")  = "SYWB031"			'�X�V��۸��і�
	rsd("UpdtTmnl")  = "WEB"				'�X�V�[����
	rsd("WorkDate")  = dWorkDate(1)			'��Ɠ�
	rsd("Term")      = sHH(1)				'���ԑ�
	rsd("OpeOrder")  = iOpeOrder(1)			'��Ə���
	rsd("HopeTerm")  = sHH(1)				'��]���ԑ�
	rsd("DualOpeNo") = iDualOpeNo(1)		'�f���A����Ɣԍ�
	rsd("SendFlag")  = "Y"					'���M�t���O
'	rsd("AppDate") 	 = AppDate(1)			'�\����
	rsd.Update
	rsd.Close
	If iDualOpeNo(0) > 0 Then	'�f���A����Ɣԍ�
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo(0))
		rsd("DualOpeNo") = CLng(sOpeNo)		'�f���A����Ɣԍ�
		rsd("UpdtTime")  = Now()			'�X�V����
		rsd("UpdtPgCd")  = "SYWB031"		'�X�V��۸��і�
		rsd("UpdtTmnl")  = "WEB"			'�X�V�[����
		rsd("SendFlag")  = "Y"				'���M�t���O
		rsd.Update
		rsd.Close
	End If
	sErrMsg = ""

End Function

%>