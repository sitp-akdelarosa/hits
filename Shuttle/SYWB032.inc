<% 

'�V���[�V�g�p�\��̃`�F�b�N
''�w��V���[�V�̎g�p�\�肪���邩�`�F�b�N����
Function ChkAppCha(conn, rsd, sChassisID)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sChassisID	[ I ]�V���[�VID
'	�߂�l		[ O ]True�F�\��Ȃ��@False�F�\�肠��

	Dim sql
	Dim sDateNow

	'���ݓ��t�擾
	sDateNow = GetYMDStr(Date())

	sql = "SELECT OpeNo FROM sAppliInfo" & _
			" WHERE ChassisId = '" & sChassisID & "'" & _
			  " AND WorkDate >= '" & sDateNow & "'"
	rsd.Open sql, conn, 0, 1, 1
	ChkAppCha = rsd.EOF	
	rsd.close

End Function

'�V���[�V�����ύX�ɂ��_�C�����ς��\�����擾
Sub GetChangeApp(conn, rsd, sChassisID, iOpeNo, iDualOpeNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sChassisID	[ I ]�V���[�VID
'	iOpeNo		[ O ]������Ɣԍ��i0�F�Y����ƂȂ��j
'	iDualOpeNo	[ O ]���o��Ɣԍ��i0�F�Y����ƂȂ��j

	Dim sql

	iOpeNo		= 0		'������Ɣԍ��i0�F�Y����ƂȂ��j
	iDualOpeNo	= 0		'���o��Ɣԍ��i0�F�Y����ƂȂ��j

	'�w��V���[�V���g�p���̔����R���e�i���擾
	sql = "SELECT OpeNo,DualOpeNo FROM sAppliInfo" & _
		  " WHERE ChassisId = '" & sChassisID & "'" & _
			" AND RecDel = 'R'" & _
			" AND Status = '02'" & _
			" AND Place = 'SY'" & _
			" AND DelFlag <> 'Y'" & _
			" AND LockFlag <> 'Y'"
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then		'�Y�����R�[�h����
		iOpeNo		= rsd("OpeNo")			'������Ɣԍ�
		iDualOpeNo	= rsd("DualOpeNo")		'���o��Ɣԍ�
	End If
	rsd.close
End Sub

%>