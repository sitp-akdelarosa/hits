<% 

'�w�莞�ԑсA�w��O���[�v�̐\�������擾
Sub GetAppHH(conn, rsd, _
				sGrpID, sTrgDate, TimeSlot, TimeNo, _
				sDateNow, iCurTime, iNextTime, iNextApp, _
				iRecCnt, iDelCnt, iCloseMode, _
				iAppCnt, _
				iAppOpeNo(), sAppUserNm(), sAppContNo(), _
				sAppBLNo(), sAppRecDel(), sAppStatus(), _
				sAppPlace(), sAppChassisId(), _
				sAppWorkFlag(), sAppCReason(), sAppContSize(), _
				sAppTerm(), sAppHopeTerm(), iAppOpeOrder(), _
				iAppDualOpeNo(), sAppFromTo(), sAppDelFlag(), sDelChaStock(), sAppTerminal(), sAppVPBookNo())
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sGrpID			[ I ]�O���[�vID
'	sTrgDate		[ I ]�Ώۓ�("YYYYMMDD")
'	TimeSlot		[ I ]���ԑыL���i��F08�`16,A,B,D�j
'	TimeNo			[ I ]���ԑєԍ��i��F8�`16,30,31,32�j
'	sDateNow		[ I ]���ݓ��t("YYYYMMDD")
'	iCurTime		[ I ]���ݍ�ƒ����ԑсi��F8�`16,31�j
'	iNextTime		[ I ]����Ǝ��ԑсi��F8�`16,31�j
'	iNextApp		[ I ]����Ɛ\�����i�T���O�m�莞�j
'	iRecCnt			[ O ]�����{��
'	iDelCnt			[ O ]���o�{��
'	iCloseMode		[ O ]�������[�h�i0�F�^�s�O�@1�F�����@2�F�^�s���@3�F�m��@4�F�m�蒆�@-1�F�J�����j
'	iAppCnt			[I/O]�\�����
'	iAppOpeNo()		[ O ]��Ɣԍ�
'	sAppUserNm()		[ O ]���[�U��
'	sAppContNo()		[ O ]�R���e�i�ԍ�
'	sAppBLNo()		[ O ]�a�k�ԍ�
'	sAppRecDel()		[ O ]���o���敪
'	sAppStatus()		[ O ]���
'	sAppPlace()		[ O ]�ꏊ
'   	sAppChassisId()		[ O ]�V���[�VID
'	sAppWorkFlag()		[ O ]��ƒ��t���O
'	sAppCReason()		[ O ]�L�����Z�����R
'	sAppContSize()		[ O ]�R���e�i�T�C�Y
'	sAppTerm()		[ O ]���ԑ�
'	sAppHopeTerm()		[ O ]��]���ԑ�
'	iAppOpeOrder()		[ O ]��Ə���
'	iAppDualOpeNo()		[ O ]�f���A����Ɣԍ�
'	sAppFromTo()		[ O ]���o��^������
'	sAppDelFlag()		[ O ]�폜�t���O
'	sDelChaStock()		[ O ]���o�w��V���[�V�̍݌Ɂi"E"�F��݌Ɂj
'	sAppTerminal()		[ O ]�^�[�~�i������
'	sAppVPBookNo()		[ O ]�u�o�w�莞�̃u�b�L���O�ԍ�

 	Dim sAppFlag		'�\�񒆂̗L���iY=����j
	Dim sNotLockFlag	'�����b�N�̗L���iY=����j
	Dim	iOOrder			'��Ə���
	Dim iWk				'���[�N

	'�\�����̎擾�i�w����A�w�莞�ԑсj
	Call GetAppHHAll(conn, rsd, sTrgDate, TimeSlot)
	
	iRecCnt = 0			'�����{��
	iDelCnt = 0			'���o�{��
	iOOrder = 1			'��Ə���
	sAppFlag = ""		'�\�񒆂̗L���iY=����j
	sNotLockFlag = ""	'�����b�N�̗L���iY=����j
	Do Until rsd.EOF
		'�S���o���{���̃J�E���g
		If Trim(rsd("RecDel")) = "R" Then	'����
			iRecCnt = iRecCnt + 1		'�����{��
		Else								'���o
			iDelCnt = iDelCnt + 1		'���o�{��
		End If
		'�\�񒆂̗L���`�F�b�N
		If Trim(rsd("Status")) = "02" Then	'���
			sAppFlag = "Y"			'�\�񒆂̗L���iY=����j
		End If
		'�����b�N�̗L���`�F�b�N
		If Trim(rsd("LockFlag")) <> "Y" Then	'���
			sNotLockFlag = "Y"		'�����b�N�̗L���iY=����j
		End If
		'��Ə��ʂ̌v�Z
		iWk = rsd("OpeOrder")		'��Ə���
		If iWk > 0 And iWk < 999 Then	'��Ə��ʂ��L��
			''���ԑї\��łȂ� OR �L�����Z�� �����O
			iWk = iOOrder
			iOOrder = iOOrder + 1	'��Ə���
		End If
		'�ΏۃO���[�v�̏ꍇ�̓��R�[�h���L��
		If Trim(rsd("GroupID")) = sGrpID Then
			iAppOpeNo(iAppCnt)      = rsd("OpeNo")				'��Ɣԍ�
			sAppUserNm(iAppCnt)     = Trim(rsd("UName"))		'���[�U��
			sAppContNo(iAppCnt)		= Trim(rsd("ContNo"))		'�R���e�i�ԍ�
			sAppBLNo(iAppCnt)		= Trim(rsd("BLNo"))			'�a�k�ԍ�
			sAppRecDel(iAppCnt)     = Trim(rsd("RecDel"))		'���o���敪
			sAppStatus(iAppCnt)     = Trim(rsd("Status"))		'���
			sAppPlace(iAppCnt)      = Trim(rsd("Place"))		'�ꏊ
			sAppChassisId(iAppCnt)  = Trim(rsd("ChassisId"))	'�V���[�VID
			sAppWorkFlag(iAppCnt)   = Trim(rsd("WorkFlag"))		'��ƒ��t���O
			sAppCReason(iAppCnt)    = Trim(rsd("CReason"))		'�L�����Z�����R
			iAppDualOpeNo(iAppCnt)  = rsd("DualOpeNo")			'�f���A����Ɣԍ�
			sAppHopeTerm(iAppCnt)  	= Trim(rsd("HopeTerm"))		'��]���ԑ�
			sAppContSize(iAppCnt)  	= Trim(rsd("ContSize"))		'�R���e�i�T�C�Y
			sAppDelFlag(iAppCnt)	= Trim(rsd("DelFlag"))		'�폜�t���O
			If sAppRecDel(iAppCnt) = "D" Or sAppRecDel(iAppCnt) = "V" Then	'���o�� Or VP�Ή�
				sAppFromTo(iAppCnt) = Trim(rsd("DeliverTo"))	'���o��
			Else								'������
				sAppFromTo(iAppCnt) = Trim(rsd("ReceiveFrom"))	'������
			End If
			sDelChaStock(iAppCnt) = ""							'���o�w��V���[�V�̍݌�

			sAppTerm(iAppCnt) 		= TimeSlot					'���ԑ�
			iAppOpeOrder(iAppCnt)	= iWk						'��Ə���
			sAppTerminal(iAppCnt)	= Trim(rsd("TName"))		'�^�[�~�i������(01/10/01)
			sAppVPBookNo(iAppCnt)	= Trim(rsd("VPBookNo"))		'VP�u�b�L���O�ԍ�(01/10/01)

			iAppCnt = iAppCnt + 1		'�\�����
		End If
		rsd.MoveNext
	Loop
	rsd.Close


	'���o�\��ŃV���[�V�w��̏ꍇ�̓V���[�V�̂r�x�݌ɂ��`�F�b�N
	If sTrgDate <> "WAIT" Then
		For iWk = 0 To iAppCnt - 1
			If sAppRecDel(iWk) = "D" and _
			   sAppDelFlag(iWk) <> "Y" and _
			   sAppWorkFlag(iWk) <> "Y" and _
		       iAppDualOpeNo(iWk) = 0 and _
			   sAppStatus(iWk) = "02" and _
		       sAppChassisId(iWk) <> "" Then
				'��V���[�V�r�x�݌ɂ̔���
				If Not ChkEmpChaStock(conn, rsd, sAppChassisId(iWk)) Then	'�Ȃ�
					sDelChaStock(iWk) = "E"	'���o�w��V���[�V�̍݌Ɂ���݌�
				End If
			End If
		Next
	End If

	'�������[�h�̌v�Z�i0�F�^�s�O�@1�F�����@2�F�^�s���@3�F�m��@4�F�m�蒆�@-1�F�J�����j
	iCloseMode = 0	'�^�s�O
	If sTrgDate = sDateNow Then			'�\�����{���̏ꍇ
		If TimeNo = 12 Then				'�ߑO�\��̏ꍇ
			If 11 < iNextTime Then		'����11���ȍ~�̏ꍇ
				iCloseMode = 1	'����
			ElseIf 11 = iNextTime Then	'����11���̏ꍇ
				If iNextStat = 2 Then	'�T���O
					iCloseMode = 1	'����
				End If
			End If
		ElseIf TimeNo = 30 Then			'�ߌ�\��̏ꍇ
			If iNextTime = 31 Then			'�����[�ς̏ꍇ
				iCloseMode = 1	'����
			ElseIf iNextTime = 15 Then		'15���̏ꍇ
				If iNextStat = 2 Then		'�T���O
					iCloseMode = 1	'����
				End If
			End If
		ElseIf TimeNo < iNextTime Then	'�^�s���ԈȑO�̏ꍇ
			If sAppFlag = "Y" Then		'�\�񒆂̗L���iY=����j
				iCloseMode = 2	'�^�s��
			Else
				iCloseMode = 1	'����
			End If
		ElseIf TimeNo = iNextTime Then	'���^�s�̏ꍇ
			If iNextStat = 1 Then		'�R�O���O
				iCloseMode = -1	'�J����
			ElseIf iNextStat = 2 Then	'�T���O
				iCloseMode = 3	'�m��
				'�����b�N�L�肩�m��{�����قȂ�ꍇ
				If sNotLockFlag = "Y" Or _
				   (iRecCnt + iDelCnt) <> iNextApp Then
					iCloseMode = 4	'�m�蒆
				End If
			End If
		End If
	ElseIf sTrgDate < sDateNow Then		'�ߋ��̏ꍇ
		iCloseMode = 1	'����
	End If

End Sub

'�\�����̎擾�i�w����A�w�莞�ԑсj
'�@���[�U�������[�U�e�[�u�����擾
Sub GetAppHHAll(conn, rsd, WorkDate, Term)
'conn		[ I ]DB�ڑ�Object
'rsd		[I/O]���R�[�h�Z�b�gObject
'WorkDate	[ I ]��Ɠ�
'Term		[ I ]���ԑсiHH�j
'�߂�l		[ O ]�Ȃ�

	Dim 	sql

	If Term = "D" Then	'���[�U�폜�̎擾
		'�w����̍폜���R�[�h�擾
		sql = "SELECT sAppliInfo.*, sMUser.Name as UName, sTerminal.Name as TName FROM sAppliInfo, sMUser, sTerminal" & _
			  " WHERE sAppliInfo.WorkDate = '" & WorkDate & "'" & _
			  " AND sAppliInfo.DelFlag = 'Y'" & _
			  " AND RTRIM(sAppliInfo.UserID) = sMUser.UserID"  & _
			  " AND RTRIM(sAppliInfo.Terminal) = sTerminal.Terminal"	'VP�Ή�
	Else				'�ʏ�̎擾
		'�w����A�w�莞�ԑт̍폜�łȂ����R�[�h�擾
		sql = "SELECT sAppliInfo.*, sMUser.Name as UName, sTerminal.Name as TName FROM sAppliInfo, sMUser, sTerminal" & _
			  " WHERE sAppliInfo.WorkDate = '" & WorkDate & "'" & _
			  " AND sAppliInfo.Term = '" & Term & "'" & _
			  " AND sAppliInfo.DelFlag <> 'Y'" & _
			  " AND RTRIM(sAppliInfo.UserID) = sMUser.UserID"  & _
			  " AND RTRIM(sAppliInfo.Terminal) = sTerminal.Terminal"	'VP�Ή�
	End If
	sql = sql & " Order By sAppliInfo.DelFlag, sAppliInfo.Term, " & _
				" sAppliInfo.OpeOrder, sAppliInfo.HopeTerm, sAppliInfo.AppDate"
	rsd.Open sql, conn, 0, 1, 1

End Sub

'�w��O���[�v�̈������҂��\�����擾
Sub GetAppWait(conn, rsd, _
				sGrpID, _
				iAppCnt, _
				iAppOpeNo, sAppUserNm, sAppContNo, _
				sAppBLNo, sAppRecDel, sAppStatus, _
				sAppPlace, sAppChassisId, _
				sAppWorkFlag, sAppCReason, sAppContSize, _
				sAppTerm, sAppHopeTerm, iAppOpeOrder, _
				iAppDualOpeNo, sAppFromTo, sAppDelFlag, sAppTerminal, sAppVPBookNo)
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sGrpID			[ I ]�O���[�vID
'	iAppCnt			[ O ]�\�����
'	iAppOpeNo()		[ O ]��Ɣԍ�
'	sAppUserNm()		[ O ]���[�U��
'	sAppContNo()		[ O ]�R���e�i�ԍ�
'	sAppBLNo()		[ O ]�a�k�ԍ�
'	sAppRecDel()		[ O ]���o���敪
'	sAppStatus()		[ O ]���
'	sAppPlace()		[ O ]�ꏊ
'   	sAppChassisId()		[ O ]�V���[�VID
'	sAppWorkFlag()		[ O ]��ƒ��t���O
'	sAppCReason()		[ O ]�L�����Z�����R
'	sAppContSize()		[ O ]�R���e�i�T�C�Y
'	sAppTerm()		[ O ]���ԑ�
'	sAppHopeTerm()		[ O ]��]���ԑ�
'	iAppOpeOrder()		[ O ]��Ə���
'	iAppDualOpeNo()		[ O ]�f���A����Ɣԍ�
'	sAppFromTo()		[ O ]���o��^������
'	sAppDelFlag()		[ O ]�폜�t���O
'	sAppTerminal()		[ O ]�^�[�~�i������
'	sAppVPBookNo()		[ O ]�u�o�w�莞�̃u�b�L���O�ԍ�

	iAppCnt = 0		'�\�����

	'�\�����̎擾�i�w����A�w��O���[�v�j
	Call GetAppInfo(conn, rsd, sGrpID, "WAIT")
	Do Until rsd.EOF
		iAppOpeNo(iAppCnt)      = rsd("OpeNo")			'��Ɣԍ�
		sAppUserNm(iAppCnt)     = Trim(rsd("UName"))		'���[�U��
		sAppContNo(iAppCnt)	= Trim(rsd("ContNo"))		'�R���e�i�ԍ�
		sAppBLNo(iAppCnt)	= Trim(rsd("BLNo"))		'�a�k�ԍ�
		sAppRecDel(iAppCnt)     = Trim(rsd("RecDel"))		'���o���敪
		sAppStatus(iAppCnt)     = Trim(rsd("Status"))		'���
		sAppPlace(iAppCnt)      = Trim(rsd("Place"))		'�ꏊ
		sAppChassisId(iAppCnt)  = Trim(rsd("ChassisId"))	'�V���[�VID
		sAppWorkFlag(iAppCnt)   = Trim(rsd("WorkFlag"))		'��ƒ��t���O
		sAppCReason(iAppCnt)    = Trim(rsd("CReason"))		'�L�����Z�����R
		iAppDualOpeNo(iAppCnt)  = rsd("DualOpeNo")		'�f���A����Ɣԍ�
		sAppHopeTerm(iAppCnt)  	= Trim(rsd("HopeTerm"))		'��]���ԑ�
		sAppContSize(iAppCnt)  	= Trim(rsd("ContSize"))		'�R���e�i�T�C�Y
		sAppDelFlag(iAppCnt)	= Trim(rsd("DelFlag"))		'�폜�t���O
		If sAppRecDel(iAppCnt) = "D" Then	'���o��
			sAppFromTo(iAppCnt) = Trim(rsd("DeliverTo"))	'���o��
		Else								'������
			sAppFromTo(iAppCnt) = Trim(rsd("ReceiveFrom"))	'������
		End If

		sAppTerminal(iAppCnt)	= Trim(rsd("TName"))		'�^�[�~�i������(01/10/01)
		sAppVPBookNo(iAppCnt)	= Trim(rsd("VPBookNo"))		'VP�u�b�L���O�ԍ�(01/10/01)

		sAppTerm(iAppCnt) 	= Trim(rsd("Term"))		'���ԑ�
		iAppOpeOrder(iAppCnt)	= 0				'��Ə���

		iAppCnt = iAppCnt + 1					'�\�����

		rsd.MoveNext
	Loop
	rsd.Close

End Sub

'�V���[�V�ݒ�
Sub SetAppChas( iAppCnt, _
				iAppOpeNo(), sAppUserNm(), sAppContNo(), _
				sAppBLNo(), sAppRecDel(), sAppStatus(), _
				sAppPlace(), sAppChassisId(), _
				sAppWorkFlag(), sAppCReason(), sAppContSize(), _
				sAppTerm(), sAppHopeTerm(), iAppOpeOrder(), _
				iAppDualOpeNo(), sAppFromTo())
'	iAppCnt		[ I ]�\�����
'	iAppOpeNo()	[ I ]��Ɣԍ�
'	sAppUserNm()	[ I ]���[�U��
'	sAppContNo()	[ I ]�R���e�i�ԍ�
'	sAppBLNo()	[ I ]�a�k�ԍ�
'	sAppRecDel()	[ I ]���o���敪
'	sAppStatus()	[ I ]���
'	sAppPlace()	[ I ]�ꏊ
'   	sAppChassisId()	[I/O]�V���[�VID
'	sAppWorkFlag()	[ I ]��ƒ��t���O
'	sAppCReason()	[ I ]�L�����Z�����R
'	sAppContSize()	[ I ]�R���e�i�T�C�Y
'	sAppTerm()	[ I ]���ԑ�
'	sAppHopeTerm()	[ I ]��]���ԑ�
'	iAppOpeOrder()	[ I ]��Ə���
'	iAppDualOpeNo()	[ I ]�f���A����Ɣԍ�
'	sAppFromTo()	[ I ]���o��^������

	Dim i, k

	'�f���A���Ŕ������V���[�V�����肵�Ă���ꍇ�ɔ��o���ɃV���[�V���Z�b�g
	For i = 0 To iAppCnt - 1
		If sAppRecDel(i) = "R" And _
		   sAppChassisId(i) <> "" And _
		   iAppDualOpeNo(i) > 0 And _
		   sAppStatus(i) = "02" Then
			'�f���A���̔��o��������
			For k = 0 To iAppCnt - 1
				If iAppOpeNo(k) = iAppDualOpeNo(i) Then
					'���o���V���[�V�ɃZ�b�g
					sAppChassisId(k) = sAppChassisId(i)
					Exit For
				End If
			Next
		End If
	Next

End Sub

'���ԑуZ���̐ݒ�
Sub SetCell01(conn, rsd, sTrgDate, TimeSlot, _
				sSTime, sETime, iSTime, iETime, _
				iCloseMode, sOpenFlag(), _
				TimeName, TimeJmp, sStatus)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sTrgDate	[ I ]�w���("YYYYMMDD")
'	TimeSlot	[ I ]���ԑыL���i��F08�`16,A,B,D�j
'	sSTime		[ I ]�J�n���ԁiHHMM�j
'	sETime		[ I ]�I�����ԁiHHMM�j
'	iSTime		[ I ]�J�n���ԑ�
'	iETime		[ I ]�I�����ԑ�
'	iCloseMode	[ I ]�������[�h�i0�F�^�s�O�@1�F�����@2�F�^�s���@3�F�m��@4�F�m�蒆�@-1�F�J�����j
'	sOpenFlag	[ I ]�J���t���O
'	TimeName	[ O ]���ԑі�
'	TimeJmp		[ O ]���ԑуW�����v
'	sStatus		[ O ]���ԑя��

	Dim iTimeNo		'���ԑєԍ��i0�`23�j

	'���ԑі��擾
	iTimeNo = -1
	If TimeSlot = "12" Then		'�ߑO�w��
		TimeName = "�ߑO�w��"
		If iCloseMode = 0 Then	'�^�s�O
			sStatus = "��t��"
		Else					'�^�s���A�^�s��
			sStatus = "����"
		End If
	ElseIf TimeSlot = "A" Then	'�ߌ�w��
		TimeName = "�ߌ�w��"
		If iCloseMode = 0 Then	'�^�s�O
			sStatus = "��t��"
		Else					'�^�s���A�^�s��
			sStatus = "����"
		End If
	ElseIf TimeSlot = "B" Then	'�[�ώw��
		TimeName = "�[�ώw��"
		sStatus = "��t��"
		If iCloseMode = 1 Then			'����
			sStatus = "����"
		ElseIf iCloseMode = 2 Then		'�^�s��
			sStatus = "�^�s��"
		ElseIf iCloseMode = 3 Then		'�m��
			sStatus = "�m��"
		ElseIf iCloseMode = 4 Then		'�m�蒆
			sStatus = "�m�蒆"
		End If
	ElseIf TimeSlot = "D" Then	'���[�U�폜
		TimeName = "���[�U�폜"
		sStatus = "�@"
	Else						'���ԑюw��
		'�\��������v�Z
		iTimeNo = CLng(TimeSlot)
		If iTimeNo < 12 Then	'�ߑO
			If iTimeNo = iSTime Then
				TimeName = GetTimeSlot(CLng(Left(sSTime, 2)), _
									   CLng(Right(sSTime, 2)), "S")
			Else
				TimeName = GetTimeSlot(iTimeNo, 0, "S")
			End If
		Else					'�ߌ�
			If iTimeNo = iETime Then
				TimeName = GetTimeSlot(CLng(Left(sETime, 2)), _
									   CLng(Right(sETime, 2)), "E")
			Else
				TimeName = GetTimeSlot(iTimeNo, 0, "S")
			End If
		End If

		'��Ԍv�Z
		'�O���[�v�g�J����ʂւ̃W�����v��ݒ�
		sStatus = "<A href=""SYWB019.asp?YMD=" & sTrgDate & TimeSlot & "&NAME=" & TimeName & " "">"
		If sOpenFlag(iTimeNo) = "Y" Then
			sStatus = sStatus & "�J���\��"
		Else
			sStatus = sStatus & "�J�����Ȃ�"
		End If

		Select Case iCloseMode
   			Case 1					'����
				sStatus = "����"
    		Case 2					'�^�s��
				sStatus = "�^�s��"
    		Case -1					'�J����
				If sOpenFlag(iTimeNo) = "Y" Then
					sStatus = "�J����"
				Else
					sStatus = "�J�����Ȃ�"
				End If
    		Case 3					'�m��
				sStatus = "�m��"
    		Case 4					'�m�蒆
				sStatus = "�m�蒆"
		End Select 
	End If

	'�\��ւ̃W�����v
	TimeJmp = ""
	if iCloseMode < 1 and TimeSlot <> "D" Then		'�������̏ꍇ
		'���ԑт���\��ւ̃W�����v��ݒ�
'VP�Ή�
'''		TimeJmp = "<A href=""SYWB010.asp?YMD=" & sTrgDate & TimeSlot "&NAME=" & TimeName & " "">"
		TimeJmp = "<A href=""SYWB011.asp?YMD=" & sTrgDate & TimeSlot & _
												"&NAME=" & TimeName & " "">"
	End If

End Sub

'�f�[�^�Z���̐ݒ�(VP�Ή�)
Sub SetCell05(iRecIdx, iCloseMode, _
				sTrgDate, TimeSlot, TimeName, _
				iAppOpeNo(),  sAppUserNm(), sAppContNo(), sAppBLNo(), _
          		sAppRecDel(), sAppStatus(),  sAppPlace(), _
          		sAppChassisId(), sAppWorkFlag(), sAppCReason(), _
          		sAppContSize(), sAppTerm(), sAppHopeTerm(), _
				iAppOpeOrder(), iAppDualOpeNo(), sAppFromTo(), _
				sAppDelFlag(), sDelChaStock(), sAppTerminal(), sAppVPBookNo(), _
				sCell())

'	iRecIdx			[ I ]���R�[�h�C���f�b�N�X
'	iCloseMode		[ I ]�������[�h
'	sTrgDate		[ I ]���t
'	TimeSlot		[ I ]���ԑ�
'	TimeName		[ I ]���ԑі�
'	iAppOpeNo		[ I ]��Ɣԍ�
'	sAppUserNm		[ I ]���[�U��
'	sAppContNo		[ I ]�R���e�i�ԍ�
'	sAppBLNo		[ I ]�a�k�ԍ�
'	sAppRecDel		[ I ]���o���敪
'	sAppStatus		[ I ]���
'	sAppPlace		[ I ]�ꏊ
'	sAppChassisId		[ I ]�V���[�VID
'	sAppWorkFlag		[ I ]��ƒ��t���O
'	sAppCReason		[ I ]�����폜���R
'	sAppContSize		[ I ]�R���e�i�T�C�Y
'	sAppTerm		[ I ]���莞�ԑсiHH�j
'	sAppHopeTerm		[ I ]�\�񎞎��ԑсiHH�j
'	iAppOpeOrder		[ I ]��Ə���
'	iAppDualOpeNo		[ I ]�f���A����Ɣԍ�
'	sAppFromTo		[ I ]���o��^������
'	sAppDelFlag		[ I ]�폜�t���O
'	sDelChaStock		[ I ]���o�w��V���[�V�̍݌Ɂi"E"�F�݌ɂȂ��j
'	sAppTerminal		[ I ]�^�[�~�i������
'	sAppVPBookNo		[ I ]�u�o�w�莞�̃u�b�L���O�ԍ�
'	sCell			[ O ]�Z�����e
	Dim sOpeNo		'��Ɣԍ�
	Dim sType		'�^�C�v
	Dim sTerm_Name	'�^�[�~�i������

	If iRecIdx = -1 Then	'�Ή����R�[�h�Ȃ����󔒍s
		sCell(0) = "�@"
		sCell(1) = "�@"
		sCell(2) = "�@"
		sCell(3) = "�@"
		sCell(4) = "�@"
		sCell(5) = "�@"
		sCell(6) = "�@"
		sCell(7) = "�@"
		sCell(8) = "�@"
		sCell(9) = "�@"
		sCell(10) = "�@"
		Exit Sub
	End If
	If iRecIdx < 0 Then		'�P�Ɣ��o���\���p�̋󔒍s
		sCell(0) = "�|"
		sCell(1) = "�|"
		sCell(2) = "�|"
		sCell(3) = "�|"
		sCell(4) = "�|"
		sCell(5) = "�|"
		sCell(6) = "�|"
		sCell(7) = "�|"
		sCell(8) = "�|"
		sCell(9) = "�|"
		sCell(10) = "�|"
		If iRecIdx = -2 Then
			sCell(5) = "��"
		Else
			sCell(5) = "�o"
		End If
		Exit Sub
	End If

	'�\�����R�[�h����
'''		'���[�U��	okimi
		'�Ώۂb�x�^�u�o
		'�b�x�^�u�o�敪�E�^�[�~�i���擾(VP�Ή�)
		sCell(0) = sAppTerminal(iRecIdx)

		'��Ɣԍ�
		sOpeNo = GetNumStr(iAppOpeNo(iRecIdx), 5)	'�Œ茅���ł̕\��
		sCell(1) = sOpeNo
		If iAppDualOpeNo(iRecIdx) > 0 Then
			sCell(1) = sCell(1) & "*"
		End If

		'��Ə���
		sCell(2) = CStr(iAppOpeOrder(iRecIdx))
		If sCell(2) = "0" Or _
		   sCell(2) = "999" Then
			sCell(2) = "-"
		End If

		'�R���e�i�ԍ��^�a�k�ԍ��^�u�b�L���O�ԍ�
		sCell(3) = ""
		If sTrgDate <> "WAIT" Then
			If TimeSlot > "B" Or _
			   sAppStatus(iRecIdx) = "03" Or _
			   iCloseMode < 1 Then
				If sAppVPBookNo(iRecIdx) = "" Then	'VP�Ή�
					sCell(3) = "<A href=""SYWB020.asp?YMD=" & sTrgDate & TimeSlot & _
					"&NAME=" & TimeName & _
					"&TNAME=" & sAppTerminal(iRecIdx) & _
					"&OPENO=" & sOpeNo & " "">"
				Else
					If sAppCReason(iRecIdx) <> "9"  then	'011101 �{���s���̎��A���J�[�Ȃ�
						sCell(3) = "<A href=""SYWB064.asp?YMD=" & sTrgDate & TimeSlot & _
						"&NAME=" & TimeName & _
						"&TNAME=" & sAppTerminal(iRecIdx) & _
						"&OPENO=" & sOpeNo & " "">"
					End If
				End If
			End If
		End If
		If sAppContNo(iRecIdx) <> "" and sAppVPBookNo(iRecIdx) = "" Then	'VP�Ή�
			sCell(3) = sCell(3) & sAppContNo(iRecIdx)
		Elseif sAppBLNo(iRecIdx) <> "" Then
			sCell(3) = sCell(3) & sAppBLNo(iRecIdx)
		Else
			sCell(3) = sCell(3) & sAppVPBookNo(iRecIdx)
		End If
		sCell(3) = sCell(3) & "</A>"

		'�\��^�C�v
		Select Case sAppHopeTerm(iRecIdx) 
    		Case "12"
				sCell(4) = "�ߑO"
    		Case "A"
				sCell(4) = "�ߌ�"
    		Case "B"
				sCell(4) = "�[��"
    		Case Else 
				sCell(4) = "����"
		End Select 

		'���^�o�^��敪(01/10/01)
		Select Case sAppRecDel(iRecIdx)
    		Case "R"
				sCell(5) = "��"
    		Case "D"
				sCell(5) = "�o"
    		Case "V"
				sCell(5) = "��"
		End Select 

'''		If sAppRecDel(iRecIdx) = "R" Then
'''			sCell(5) = "��"
'''		Else
'''			sCell(5) = "�o"
'''		End If

		'�R���e�i�T�C�Y
		sCell(6) = sAppContSize(iRecIdx)

		'�V���[�VID
		If sTrgDate <> "WAIT" Then
			'�����̏ꍇ�̓V���[�V�����ύX�ďo���\
			If sAppChassisId(iRecIdx) = "" Then
				sCell(7) = "-"
			Else
				'�V���[�V������ꍇ�̂�
				sCell(7) = "<A href=""SYWB024.asp?YMD=" & sTrgDate & _
												"&TRGID=" & sAppChassisId(iRecIdx) & _
												" "">"
			End If

			If sAppRecDel(iRecIdx) = "D" Then
				'���o�̏ꍇ�͎g�p�V���[�V�ύX�ďo���\
				If iCloseMode = 0 And _
					TimeSlot <> "D" Then
					sCell(7) = "<A href=""SYWB022.asp?YMD=" & sTrgDate & _
												"&OPENO=" & sOpeNo & _
												"&CID=" & sAppChassisId(iRecIdx) & _
												" "">"
					If sAppChassisId(iRecIdx) = "" Then
						sCell(7) = sCell(7) & "-"
					End If
				End If
			End If

			sCell(7) = sCell(7) & sAppChassisId(iRecIdx) & "</A>"
		Else
			If sAppChassisId(iRecIdx) = "" Then
				sCell(7) = "�@"
			Else
				sCell(7) = sAppChassisId(iRecIdx)
			End If
		End If

		'�ꏊ
		If sAppPlace(iRecIdx) = "" Then
			sCell(8) = "-"
		ElseIf sAppPlace(iRecIdx) = "MV" Then
			sCell(8) = "-"
		Else
'VP�Ή����ŁE����b�x�̎��͂b�x��\������
'2003/08/25 ICCT
'''''		If sAppPlace(iRecIdx) = "VP" and (sAppTerminal(iRecIdx) = "���łb�x" or sAppTerminal(iRecIdx) = "����b�x")Then 
			If sAppPlace(iRecIdx) = "VP" and (sAppTerminal(iRecIdx) = "���łb�x" or sAppTerminal(iRecIdx) = "�h�b�b�s")Then 
				sCell(8) = "CY"
			Else
				sCell(8) = sAppPlace(iRecIdx)
			End If
		End If

		'��ԁA���l
		sCell(9) = "�@"
		sCell(10) = "�@"

		'���o�\��Ŏw��V���[�V���r�x�݌ɂłȂ��ꍇ
		If sDelChaStock(iRecIdx) = "E" Then
			sCell(10) = "������"
		End If

		If sAppDelFlag(iRecIdx) = "Y" Then		'���[�U�폜
			'�폜��
			sCell(9) = "�폜"
			If sAppPlace(iRecIdx) = "SY" Then
				sCell(10) = "�����"
			End If
		Else
			Select Case sAppStatus(iRecIdx) 
    			Case "02"	'������
					If sAppWorkFlag(iRecIdx) = "Y" Then
						sCell(9) = "��ƒ�"
					Else
						If sAppFromTo(iRecIdx) = "" Then
							If sAppRecDel(iRecIdx) = "R" Then
								sCell(10) = "��������"
'							ElseIF sAppRecDel(iRecIdx) = "D" Then	'VP�Ή�
							ElseIF sAppRecDel(iRecIdx) = "D" Or sAppRecDel(iRecIdx) = "V" Then	'VP�Ή�
								sCell(10) = "���o�斳"
							End If
						End If
						If sAppRecDel(iRecIdx) = "R" and _
						   sAppPlace(iRecIdx) <> "SY" Then
							sCell(10) = "����"
						End If
						sCell(9) = "��"
						Select Case sAppTerm(iRecIdx) 
    						Case "12", "A"
								sCell(9) = "����"
    						Case "B"
								'����ł�����҂��ł�����
						End Select 
					End If
    			Case "03"	'�L�����Z��
					sCell(9) = "�폜"
					Select Case sAppCReason(iRecIdx) 
    					Case "1"	'����
							sCell(10) = "����"
    					Case "2"	'BL�Ȃ�
							sCell(10) = "BL�Ȃ�"
    					Case "3"	'�V���[�V�s��
							sCell(10) = "�����s��"
    					Case "4"	'�w��V���[�V���Ȃ�
							sCell(10) = "������"
    					Case "5"	'�Q�[�g����
							If sAppRecDel(iRecIdx) = "R" Then
								sCell(10) = "�����s��"
							Elseif sAppRecDel(iRecIdx) = "D" Then	'VP�Ή�
								sCell(10) = "���o���s"
							End If
    					Case "6"	'��݌�
							sCell(10) = "��݌�"
    					Case "7"	'���ԕs��
							sCell(10) = "�\�͒�"
    					Case "8"	'���ԕs��
							If sAppRecDel(iRecIdx) = "R" Then
								sCell(10) = "��������"
'							ElseIF sAppRecDel(iRecIdx) = "D" Then	'VP�Ή�
							ElseIF sAppRecDel(iRecIdx) = "D" Or sAppRecDel(iRecIdx) = "V" Then	'VP�Ή�
								sCell(10) = "���o�斳"
							End If
    					Case "9"	'�{���s��
							sCell(10) = "�{���s��"
    					Case "A"	'�s�b�N���s
							sCell(10) = "�߯����s"
    					Case "B"	'�[�ςݕs��
							sCell(10) = "�[�ςݕs��"
    					Case Else 
							sCell(10) = "�s��"
					End Select 
					If sAppRecDel(iRecIdx) = "D" and _
					   sAppPlace(iRecIdx) = "SY" Then
						sCell(10) = "�߂�"
					End If
    			Case "04"	'����
					sCell(9) = "����"			
'����ҏ����ύX(�����ȊO�ŁA�ꏊ���r�x�̂���)
					If sAppRecDel(iRecIdx) <> "R" and _
					   sAppPlace(iRecIdx) = "SY" Then
						sCell(10) = "�����"
					End If
'VP�Ή��@���������ꍇ�R���e�i�ԍ���\������(MES�Ɏ��� 01/10/12)		'���l�ɕ\������
					If sAppRecDel(iRecIdx) = "V" and _
					   sAppContNo(iRecIdx) <> "" Then
						sCell(10) = Left(sAppContNo(iRecIdx), 4)  & "<br>" & mid(sAppContNo(iRecIdx), 5)
'''						sCell(3) = sAppContNo(iRecIdx)
					End If
    			Case Else 
					sCell(9) = "�s��"
			End Select 
		End If

End Sub

'��ƒ��\���̗L���`�F�b�N
Function ChkWorkingApp(conn, rsd, WorkDate, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	WorkDate	[ I ]��Ɠ�
'	sHH			[ I ]���ԑсiHH�j
'	�߂�l		[ O ]True�F����@False�F�Ȃ�

	ChkWorkingApp = False

	'�\�����̎擾�i�w����A�w�莞�ԑсA�w��O���[�v�j
	Call GetAppInfoTerm(conn, rsd, WorkDate, sHH, "")
	Do Until rsd.EOF
		If Trim(rsd("Status")) = "02" Or _
		   Trim(rsd("WorkFlag")) = "Y" Then
			ChkWorkingApp = True			
		End If
		rsd.MoveNext
	Loop
	rsd.Close
End Function

'�f�[�^�Z���J���[�̌v�Z
Sub	CalcDataColor(sColor, sCell())
'	sColor		[ O ]�f�[�^�Z���J���[
'	sCell()		[ I ]�Z���̒l

	sColor = ""
'''	If sCell(10) = "�����" �@Then			'01/12/04
	If sCell(10) = "�����" or  (sCell(5) = "��" and sCell(8) = "SY" and sCell(9) = "����") Then
		sColor = "bgcolor=""#f0f0ff"" "
	End If
	If sCell(10)= "����" or _
	   sCell(10)= "��������" or _
	   sCell(10)= "���o�斳" or _
	   sCell(10)= "������" then
		sColor = "bgcolor=""#FFFF00"" "
	End If
	If sCell(9) = "�폜" Then
		sColor = "bgcolor=""#dda0dd"" "
	End If
	If sCell(9) = "��ƒ�" Then
		sColor = "bgcolor=""#F08080"" "
	End If
End Sub

'���V���[�V�v�Z
''�R�O���O�ɒP�Ɣ��o�̃V���[�V�����ɕ\������
Sub CalcAppChas(conn, rsd, _
				sGrpID, sTrgDate, _
				iCurTime, iNextTime, iNextStat, _
				iAppCnt, _
				iAppOpeNo(), sAppUserNm(), sAppContNo(), _
				sAppBLNo(), sAppRecDel(), sAppStatus(), _
				sAppPlace(), sAppChassisId(), _
				sAppWorkFlag(), sAppCReason(), sAppContSize(), _
				sAppTerm(), sAppHopeTerm(), iAppOpeOrder(), _
				iAppDualOpeNo(), sAppFromTo())
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sGrpID			[ I ]�O���[�vID
'	sTrgDate		[ I ]�Ώۓ�("YYYYMMDD")
'	iCurTime		[ I ]����Ǝ��ԑсi��F8�`16,31�j
'	iNextTime		[ I ]����Ǝ��ԑсi��F8�`16,31�j
'	iNextStat		[ I ]����Ə󋵁i1�F�R�O���O�@2�F�T���O�j
'	iAppCnt			[ I ]�\�����
'	iAppOpeNo()		[ I ]��Ɣԍ�
'	sAppUserNm()	[ I ]���[�U��
'	sAppContNo()	[ I ]�R���e�i�ԍ�
'	sAppBLNo()		[ I ]�a�k�ԍ�
'	sAppRecDel()	[ I ]���o���敪
'	sAppStatus()	[ I ]���
'	sAppPlace()		[ I ]�ꏊ
'   sAppChassisId()	[I/O]�V���[�VID
'	sAppWorkFlag()	[ I ]��ƒ��t���O
'	sAppCReason()	[ I ]�L�����Z�����R
'	sAppContSize()	[ I ]�R���e�i�T�C�Y
'	sAppTerm()		[ I ]���ԑ�
'	sAppHopeTerm()	[ I ]��]���ԑ�
'	iAppOpeOrder()	[ I ]��Ə���
'	iAppDualOpeNo()	[ I ]�f���A����Ɣԍ�
'	sAppFromTo()	[ I ]���o��^������

	Dim iCCnt				'�V���[�V��
	Dim sChassisId(300)		'�V���[�VID
	Dim sSize20Flag(300)	'20�t�B�[�g�t���O
	Dim sMixSizeFlag(300)	'20/40���p�V���[�V
	Dim sNightFlag(300)		'�[�σt���O
	Dim i, k
	Dim sCurTime			'����Ǝ��ԑ�
	Dim sNextTime			'����Ǝ��ԑ�
	Dim sWk

	'���ԑєԍ����玞�ԑыL�����擾
	sCurTime  = TermNumToSym(iCurTime)
	If iNextStat > 0 Then	'�R�O���O�ȍ~�̏ꍇ�̂ݑΏۂƂ���
		sNextTime = TermNumToSym(iNextTime)
	Else
		sNextTime = ""
	End If

	'�g�p�\��V���[�V�̎擾
	Call GetEmpCha(conn, rsd, sGrpID, sTrgDate, _
				iCCnt, sChassisId, sSize20Flag, sMixSizeFlag, sNightFlag)

	For i = 0 To iAppCnt - 1
		'�P�Ɣ��o�ŃV���[�V�w�肪�Ȃ��\���ɑ΂��Čv�Z
'''		If sAppRecDel(i) = "D" And _
		If (sAppRecDel(i) = "D" OR sAppRecDel(i) = "V") And _
		   sAppChassisId(i) = "" And _
		   iAppDualOpeNo(i) = 0 And _
		   sAppStatus(i) = "02" And _
		   (sAppTerm(i) = sCurTime Or sAppTerm(i) = sNextTime) Then
			If sAppContSize(i) = "20" Then	'�\�����Q�O�t�B�[�g�̏ꍇ
				sWk = "Y"
			Else							'�\�����S�O�t�B�[�g�̏ꍇ
				sWk = ""
			End If
			'�V���[�V������
			For k = 0 To iCCnt - 1
				If sChassisId(k) <> "" Then	'�L���ȃV���[�V
					'�T�C�Y���K�����邩���p�V���[�V
					If sSize20Flag(k) = sWk Or sMixSizeFlag(k) = "Y" Then
						'�[�ϗ\��̂ݗ[�σV���[�V���g�p�\
						If sAppTerm(i) = "B" Or sNightFlag(k) <> "Y" Then
							sAppChassisId(i) = sChassisId(k)
							sChassisId(k) = ""
							Exit For
						End If
					End If
				End If
			Next
		End If
	Next
End Sub

'�g�p�\��V���[�V�̎擾
Sub GetEmpCha(conn, rsd, sGrpID, sTrgDate, _
				iCCnt, sChassisId(), _
				sSize20Flag(), sMixSizeFlag(), sNightFlag())
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sGrpID			[ I ]�O���[�vID
'	sTrgDate		[ I ]�Ώۓ�("YYYYMMDD")
'	iCCnt			[ O ]�V���[�V��
'	sChassisId()	[ O ]�V���[�VID
'	sSize20Flag()	[ O ]20�t�B�[�g�t���O
'	sMixSizeFlag()	[ O ]20/40���p�V���[�V
'	sNightFlag()	[ O ]�[�σt���O

	Dim sql
	Dim i, sWk

	'SELECT���쐬
	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
		    " AND StackFlag = 'Y'" & _
		    " AND ContFlag <> 'Y'" & _
		    " AND NotDelFlag <> 'Y'" & _
		  " Order By ChassisId"

	'����
	rsd.Open sql, conn, 0, 1, 1

	'���R�[�h�L��
	iCCnt = 0
	Do Until rsd.EOF
		sChassisId(iCCnt)   = Trim(rsd("ChassisId"))	'�V���[�VID
		sSize20Flag(iCCnt)  = Trim(rsd("Size20Flag"))	'20�t�B�[�g�t���O
		sMixSizeFlag(iCCnt) = Trim(rsd("MixSizeFlag"))	'20/40���p�V���[�V
		sNightFlag(iCCnt)   = Trim(rsd("NightFlag"))	'�[�σt���O
		iCCnt = iCCnt + 1	
		rsd.MoveNext
	Loop
	rsd.Close

	'�\�����̎擾�i�w����A�w��O���[�v�j
	Call GetAppInfoDG(conn, rsd, sGrpID, sTrgDate)
	'�{���̐\���ɂ���V���[�V�͌�₩�珜�O����
	Do Until rsd.EOF
		sWk = Trim(rsd("ChassisId"))			'�V���[�VID
		For i = 0 To iCCnt - 1
			If sChassisId(i) = sWk Then
				sChassisId(i) = ""
			End If
		Next
		rsd.MoveNext
	Loop
	rsd.Close

End Sub

'��V���[�V�r�x�݌ɂ̔���
Function ChkEmpChaStock(conn, rsd, sChassisId)
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sChassisId		[ I ]�V���[�VID
'	ReturnCode		[ O ]True�F�V���[�V����@False�F�V���[�V�Ȃ�
	Dim sql

	'SELECT���쐬
	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(ChassisId) = '" & sChassisId & "'" & _
		    " AND StackFlag = 'Y'" & _
		    " AND ContFlag <> 'Y'"

	'����
	rsd.Open sql, conn, 0, 1, 1

	ChkEmpChaStock = Not rsd.EOF
	rsd.Close

End Function

%>