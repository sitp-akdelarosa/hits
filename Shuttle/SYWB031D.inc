<% 

'�\��ύX����
Function UpdOpeUpd(conn, rsd, sOpeNo, sSend, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sOpeNo			[ I ]��Ɣԍ�
'	sSend			[ I ]�������E���o��
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
	sRDType    = Trim(rsd("RecDel"))		'���o���敪
	rsd.Close

	If sDelFlag = "Y" Then	'�폜�̏ꍇ
		If sRDType = "R" and sPlace = "SY" Then
			sErrMsg = "�Ώۂ͖߂��Ƃł�"
			Exit Function
		End If
		sErrMsg = ""
		Exit Function
	End If

	If sStatus = "04" Then	'�����̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��������Ă��܂�"
		Exit Function
	End If

	If sStatus = "03" Then	'�����̏ꍇ
		sErrMsg = "�Ώۂ͍�Ƃ��L�����Z�����Ă��܂�"
		Exit Function
	End If

	'�X�V����
	Call GetAppInfoOpeNoUpd(conn, rsd, CLng(sOpeNo))
	If sRDType = "D" Or sRDType = "V" Then	'<!--��o���Ή� -->
		rsd("DeliverTo") = sSend			'���o��
	Else
		rsd("ReceiveFrom") = sSend			'������
	End If

	rsd("SendFlag") = "Y"
	rsd("UpdtTime") = Now()				'�X�V����
	rsd("UpdtPgCd") = "SYWB031"			'�X�V��۸��і�
	rsd("UpdtTmnl") = "WEB"				'�X�V�[����
	rsd.Update
	rsd.Close
	sErrMsg = ""

End Function

%>