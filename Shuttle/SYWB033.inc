<% 
'�ύX���\�����̎擾
Sub GetOApp(conn, rsd, iOpeNo, sYMD, sErr_msg)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ O ]���R�[�h�Z�b�gObject
'	iOpeNo			[ I ]��Ɣԍ�
'	sYMD			[ I ]�Ώۓ�("YYYYMMDD")
'	sErr_msg		[ O ]�G���[���b�Z�[�W
	Dim sql

	sErr_msg = ""

	'�ύX��
	sql = "SELECT * FROM sAppliInfo" & _
			" WHERE OpeNo = " & CStr(iOpeNo) & _
			  " AND WorkDate = '" & cdate(ChgYMDStr(sYMD)) & "'"
	rsd.Open sql, conn, 0, 2, 1		'�X�V���[�h�ŃI�[�v��
	If rsd.EOF Then	
		sErr_msg = CStr(iOpeNo) & "�@�̍�Ɣԍ��͑��݂��܂���B"
	End If
	If rsd("Status") = "04" Then
		sErr_msg = CStr(iOpeNo) & "�@�̍�Ƃ͊������Ă��܂��B"
	End If
	If rsd("Status") = "03" Or rsd("DelFlag") = "Y" Then
		sErr_msg = CStr(iOpeNo) & "�@�̍�Ƃ͍폜����Ă��܂��B"
	End If
	If rsd("Place") = "SY" Then
		sErr_msg = CStr(iOpeNo) & "�@�̍�Ƃ͂r�x�ɂ���܂��B"
	End If
	If rsd("WorkFlag") = "Y" Or rsd("LockFlag") = "Y" Then 
		sErr_msg = CStr(iOpeNo) & "�@�̍�Ƃ͊m�肵�Ă���וύX�ł��܂���B"
	End If
End Sub

%>