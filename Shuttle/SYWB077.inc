<% 

'�󂫃V���[�V���̎擾
'	�w��O���[�v�E��Ɠ��E���ԑтł̍�Ɗ������_�ł�
'	�󂫃V���[�V�����v�Z����
Function GetEmptyChassisCnt(conn, rsd, sGrpID, sYMD, sHHName, i20, i40)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sGrpID		[ I ]�O���[�vID
'	sYMD		[ I ]�w����t(YYYYMMDD)
'	sHHName		[ I ]�w�莞�ԑсi"HH"�j
'	i20			[ O ]�Q�O�t�B�[�g�p��V���[�V��
'	i40			[ O ]�S�O�t�B�[�g�p��V���[�V��
'	�߂�l		[ O ]True:�v�Z�n�j False:�v�Z�s��

	Dim sql
	Dim iUC(1)		'�[�ϋ�V���[�V���Q�O�e�C�S�O�e�i�[�ςɂ̂ݎg�p�\�j
	Dim iNC(1)		'�ʏ��V���[�V���Q�O�e�C�S�O�e�i���ԑтɂ��[�ςɂ��g�p�\�j
	Dim sWkYmd, dtWk	'���[�N���t
	Dim iWk

	GetEmptyChassisCnt = False

	'�O���[�v�̌��݋󂫃V���[�V���擾
	'������
	iNC(0) = 0		'�ʏ�Q�O���i���ԑтɂ��[�ςɂ��g�p�\�j
	iNC(1) = 0		'�ʏ�S�O���i���ԑтɂ��[�ςɂ��g�p�\�j
	iUC(0) = 0		'�[�ςQ�O���i�[�ςɂ̂ݎg�p�\�j
	iUC(1) = 0		'�[�ςS�O���i�[�ςɂ̂ݎg�p�\�j
	'�V���[�V����
	'�r�x���u�����R���e�i�Ȃ������o�Ɏg���違�w��O���[�v
	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
		    " AND StackFlag = 'Y'" & _
		    " AND ContFlag <> 'Y'" & _
		    " AND NotDelFlag <> 'Y'"
	rsd.Open sql, conn, 0, 1, 1
	'�V���[�V���W�v
	Do Until rsd.EOF
		If rsd("Size20Flag") = "Y" Then
			iWk = 0		'�Q�O�e
		Else
			iWk = 1		'�S�O�e
		End If
		If rsd("NightFlag") <> "Y" Then	'�[�ςłȂ�
			iNC(iWk) = iNC(iWk) + 1	'�ʏ�i���ԑтɂ��[�ςɂ��g�p�\�j
		Else
			iUC(iWk) = iUC(iWk) + 1	'�[�ρi�[�ςɂ̂ݎg�p�\�j
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	sWkYmd = GetYMDStr(Date())	'�{���擾
	Do While sWkYmd <= sYMD		'�{������w����t�܂Ōv�Z
		'���[�N���t�̐\�������擾
		'�\����񌟍��r�p�k�쐬
		sql = "SELECT * FROM sAppliInfo" & _
				" WHERE Status = '02'" & _
				  " AND DelFlag <> 'Y'" & _
				  " AND WorkFlag = ' '" & _
				  " AND NotDelFlag <> 'Y'" & _
				  " AND DualOpeNo = 0" & _
				  " AND RTRIM(GroupID) = '" & sGrpID & "'" & _
				  " AND WorkDate = '" & sWkYmd & "'"
		If sWkYmd = sYMD Then
			'�w����̏ꍇ�͎��ԑт������ɒǉ�
			sql = sql & " AND Term <= '" & sHHName & "'"
		End If
		sql = sql & " ORDER BY Term, OpeOrder"

		'�\����񌟍�
		rsd.Open sql, conn, 0, 1, 1

		'�g�p�䐔�A�ǉ��䐔���v�Z
		Do Until rsd.EOF
			If rsd("Size20Flag") = "Y" Then
				iWk = 0		'�Q�O�e
			Else
				iWk = 1		'�S�O�e
			End If
'''			IF rsd("RecDel") = "D" Then	'���o�i�g�p�����j(VP�Ή�)
			IF rsd("RecDel") = "D" OR rsd("RecDel") = "V" Then	'���oOR��o���i�g�p�����j
				IF rsd("Term") = "B" Then	'�[��
					'�{���͗[�ς̏ꍇ���ۂɂ͂ǂ���̃V���[�V���g�p����邩������Ȃ�
					iUC(iWk) = iUC(iWk) - 1	'�[�ρi�[�ςɂ̂ݎg�p�\�j
				Else
					iNC(iWk) = iNC(iWk) - 1	'�ʏ�i���ԑтɂ��[�ςɂ��g�p�\�j
				End If
			Else						'�����i��V���[�V��������j
				IF rsd("NightFlag") = "Y" Then	'�[�ςɂ̂ݎg�p�\
					iUC(iWk) = iUC(iWk) + 1	'�[�ρi�[�ςɂ̂ݎg�p�\�j
				Else
					iNC(iWk) = iNC(iWk) + 1	'�ʏ�i���ԑтɂ��[�ςɂ��g�p�\�j
				End If
			End If

			rsd.MoveNext
		Loop
	
		rsd.Close

		'�v�Z���t�̉��Z
		dtWk = DateAdd("d", 1, CDate(ChgYMDStr(sWkYmd)))
		sWkYmd = GetYMDStr(dtWk)	'���̓��擾
	Loop

	i20 = iUC(0) + iNC(0)
	i40 = iUC(1) + iNC(1)

	GetEmptyChassisCnt = True

End Function


%>