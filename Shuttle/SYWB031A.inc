<% 

'�\��폜����
Function UpdOpeDel(conn, rsd, sOpeNo, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sOpeNo			[ I ]��Ɣԍ�
'	sErrMsg			[ O ]�G���[���b�Z�[�W

	Dim sStatus, sDelFlag, sLockFlag, iDualOpeNo, sRDType, sPlace

	sErrMsg = "�c�a�X�V�G���["

	'�\�����̎擾�i�w���Ɣԍ��j
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	If rsd.EOF Then
		rsd.Close
		sErrMsg = "�Ώۂ��폜����Ă��܂�"
		Exit Function
	End If
	sStatus    = Trim(rsd("Status"))		'���
	sDelFlag   = Trim(rsd("DelFlag"))		'�폜�t���O
	sLockFlag  = Trim(rsd("LockFlag"))		'���b�N�t���O
	iDualOpeNo =      rsd("DualOpeNo")		'�f���A����Ɣԍ�
	sRDType    = Trim(rsd("RecDel"))		'���o���敪
	sPlace     = Trim(rsd("Place"))			'�ꏊ
	rsd.Close

	If sDelFlag = "Y" Then	'�폜�̏ꍇ
		sErrMsg = "���ɍ폜����Ă��܂�"
		Exit Function
		'�ȉ��̊��S�폜�����͎g�p���Ȃ�
		If sRDType = "R" and sPlace = "SY" Then
			sErrMsg = "�Ώۂ͖߂��Ƃł�"
			Exit Function
		End If
		'���S�ɍ폜
		Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
		rsd.Delete
		rsd.Update
		rsd.Close
		sErrMsg = ""
		Exit Function
	End If

	If sStatus = "04" Then	'�����̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��������Ă��܂�"
		Exit Function
	End If

	If sStatus = "03" Then	'�L�����Z�����̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��L�����Z�����Ă��܂�"
		Exit Function
	End If

	'�ȉ��\�񒆂̏ꍇ�i�폜��Ԃɂ���j
	If sLockFlag = "Y" Then	'���b�N
		sErrMsg = "�Ώۂ͍�Ƃ��m�肵�Ă��܂�"
		Exit Function
	End If
	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	rsd("SendFlag") = "Y"
	rsd("DelFlag") = "Y"
	rsd("DualOpeNo") = 0
	rsd("OpeOrder") = 0
	rsd("UpdtTime") = Now()				'�X�V����
	rsd("UpdtPgCd") = "SYWB031"			'�X�V��۸��і�
	rsd("UpdtTmnl") = "WEB"				'�X�V�[����
	rsd.Update
	rsd.Close
	If iDualOpeNo > 0 Then	'�f���A����Ɣԍ�
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
		rsd("SendFlag") = "Y"
		rsd("DualOpeNo") = 0
		rsd("UpdtTime") = Now()			'�X�V����
		rsd("UpdtPgCd") = "SYWB031"		'�X�V��۸��і�
		rsd("UpdtTmnl") = "WEB"			'�X�V�[����
		rsd.Update
		rsd.Close
	End If
	sErrMsg = ""

End Function

%>