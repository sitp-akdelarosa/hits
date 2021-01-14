<% 
'�c�a�ڑ�
Function ConnectSvr(conn, rsd) 
	'conn		[ O ]DB�ڑ�Object
	'rsd		[ O ]���R�[�h�Z�b�gObject

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "hakataDB", "sa", ""

	Set rsd = Server.CreateObject("ADODB.Recordset")

End Function

'���[�U���̎擾
Function GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	'conn		[ I ]DB�ڑ�Object
	'rsd		[ I ]���R�[�h�Z�b�gObject
	'sUsrID		[ O ]���[�UID
	'sGrpID		[ O ]�O���[�vID
	'sUsrName	[ O ]���[�U��
	'sGrpName	[ O ]�O���[�v��
	'sOperator	[ O ]�I�y���[�^�R�[�h

	Dim sql
	Dim iWk

	'�F�؃��[�U�擾
'	sUsrID = UCase(Trim(Request.ServerVariables("AUTH_USER")))  'Commented by Seiko-denki 2003.7.26
	sUsrID = Session.Contents("userid")  'Added by Seiko-denki 2003.7.26

	If sUsrID <> "" Then
		iWk	= InStrRev(sUsrID, "\")
		If iWk > 0 Then
			sUsrID = Mid(sUsrID, iWk + 1)
		End If
	End If
	sUsrID = Left(sUsrID, 8)

	sGrpID    = ""
	sUsrName  = "�s���i" & sUsrID & "�j"
	sGrpName  = "�s��"
	sOperator = ""

	'���[�U�擾
	sql = "SELECT sMUser.Name,sMGroup.GroupName,sMGroup.GroupID,sMGroup.Operator" & _
		  " FROM sMUser,sMGroup,sMUserGroup" & _
		  " WHERE RTRIM(sMUser.UserID) = '" & sUsrID & "'" & _
		  "   AND RTRIM(sMUserGroup.UserID) = RTRIM(sMUser.UserID)" & _
		  "   AND RTRIM(sMGroup.GroupID) = RTRIM(sMUserGroup.GroupID)"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		sGrpID    = Trim(rsd("GroupID"))
		sUsrName  = Trim(rsd("Name"))
		sGrpName  = Trim(rsd("GroupName"))
		sOperator = Trim(rsd("Operator"))
	End If

	rsd.Close

End Function

'���ϐ��̎擾
Function GetEnv(conn, rsd, sEnvName)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sEnvName	[ I ]���ϐ���
'	�߂�l		[ O ]���ϐ��l

	Dim sql

	'���ϐ��擾
	sql = "SELECT EnvValue FROM sEnv" & _
		  " WHERE RTRIM(EnvName) = '" & sEnvName & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		GetEnv = Trim(rsd("EnvValue"))
	Else
		GetEnv = ""
	End If

	rsd.Close

End Function

'���t������̎擾
Function GetYMDStr(DateVal)
	'DateVal	[ I ]���t�iDate�^�j
	'�߂�l		[ O ]���t������iYYYYMMDD�`���j

	GetYMDStr = _
		Year(DateVal) & _
		Right("0" & Month(DateVal), 2) & _
		Right("0" & Day(DateVal), 2)

End Function

'���t������̕ϊ�
Function ChgYMDStr(sDate)
	'sDate		[ I ]���t������iYYYYMMDD�`���j
	'�߂�l		[ O ]���t������iYYYY/MM/DD�`���j

	ChgYMDStr = Left(sDate, 4) & "/" & _
				Mid(sDate, 5, 2) & "/" & _
				Mid(sDate, 7, 2)

End Function

'���t������̕ϊ��Q
Function ChgYMDStr2(sDate)
	'sDate		[ I ]���t������iYYYYMMDD�`���j
	'�߂�l		[ O ]���t������iYYYY�NMM��DD�`���j

	ChgYMDStr2 = Left(sDate, 4) & "�N" & _
			 	 CStr(CLng(Mid(sDate, 5, 2))) & "��" & _
				 CStr(CLng(Mid(sDate, 7, 2))) & "��"

End Function

'���t������̕ϊ��R
Function ChgYMDStr3(sDate)
	'sDate		[ I ]���t������iYYYYMMDD�`���j
	'�߂�l		[ O ]���t������iYYYY/MM/DD�`���j

	Dim sYobi
	ChgYMDStr3 = Left(sDate, 4) & "/" & _
				Mid(sDate, 5, 2) & "/" & _
				Mid(sDate, 7, 2)
	sYobi = sWeekday(Weekday(cDate(ChgYMDStr3)))
	ChgYMDStr3 = ChgYMDStr3 & "(" & sYobi & ")"
End Function

'���������t�ɕϊ�
Function ChgYMDDate(sYMD)
'	sYMD		[ I ]���t������iYYYYMMDD�`���j
'	�߂�l		[ O ]���t�iDate�^�j

	ChgYMDDate = CDate(ChgYMDStr(sYMD))
End Function

'���Ԃ����ԑтɕϊ�
Function GetTimeSlot(iH, iM, sSE)
	'iH			[ I ]����
	'iM			[ I ]��
	'sSE		[ I ]S:�J�n�AE:�I��
	'�߂�l		[ O ]���ԑѕ�����

	Dim sStr

	sStr = CStr(iH) & ":" & Right("0" & CStr(iM), 2)

	If sSE = "S" Then
		sStr = sStr & "�`" & CStr(iH + 1) & ":00"
	Else
		sStr = CStr(iH - 1) & ":00�`" & sStr
	End If

	GetTimeSlot = sStr

End Function

'���Ԃ����ԑѕ����ɕϊ�
Function GetTimeSlotStr(conn, rsd, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sHH			[ I ]���ԁi"HH"�A"A"�A"B"�j
'	�߂�l		[ O ]���ԑѕ�����

	Dim sStr, sShtStart, sShtEnd, iSTime, iETime, iHH

	Select Case Trim(sHH)
 		Case "12"
			sStr = "�ߑO�w��"
 		Case "A"
			sStr = "�ߌ�w��"
   		Case "B"
			sStr = "�[�ώw��"
    	Case Else 
			'�J�n�^�I�����Ԃ��擾
			sShtStart = GetEnv(conn, rsd, "ShtStart")
			sShtEnd   = GetEnv(conn, rsd, "ShtEnd")
			iSTime = CLng(Left(sShtStart, 2))
			iETime = CLng(Left(sShtEnd, 2))
			if Right(sShtEnd, 2) = "00" Then
'''2001/03/12			iETime = iETime - 1
				If iETime <> "16" Then
					iETime = iETime - 1
				End If
			End If

			iHH = CLng(sHH)
			If iHH = iSTime Then
				sStr = CStr(iHH) & ":" & Right(sShtStart, 2)
			Else
				sStr = CStr(iHH) & ":00"
			End If
			sStr = sStr & "�`"
			If iHH = iETime Then
				sStr = sStr & CStr(iHH) & ":" & Right(sShtEnd, 2)
			Else
				sStr = sStr & CStr(iHH + 1) & ":00"
			End If
	End Select 

	GetTimeSlotStr = sStr

End Function

'�O���[�v���ԑя��̎擾
Function GetGrpSlot(conn, rsd, sGroupID, sDate, sOpenFlag())
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sGroupID	[ I ]�O���[�v�R�[�h
'	sDate		[ I ]���t(YYYYMMDD)
'	sOpenFlag	[ O ]�J���t���O
'	�߂�l		[ O ]�Ȃ�

	Dim sql
	Dim i

	'�O���[�v�X���b�g�擾
	sql = "SELECT * FROM sGrpSlot" & _
		  " WHERE RTRIM(GroupID) = '" & sGroupID & "'" & _
		  "   AND Date = '" & sDate & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		For i = 0 To 23
			sOpenFlag(i) = Trim(rsd("OpenFlag" & Trim(Cstr(i))))
		Next
	Else
		For i = 0 To 23
			sOpenFlag(i) = "Y"
		Next
	End If

	rsd.Close

End Function

'�O���[�v���ԑя��̎擾
Function GetGrpSlotHH(conn, rsd, sGroupID, sDate, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sGroupID	[ I ]�O���[�v�R�[�h
'	sDate		[ I ]���t(YYYYMMDD)
'	sHH			[ I ]���ԑсiHH�j
'	�߂�l		[ O ]�J���t���O

	Dim sql

	'�O���[�v�X���b�g�擾
	sql = "SELECT * FROM sGrpSlot" & _
		  " WHERE RTRIM(GroupID) = '" & sGroupID & "'" & _
		  "   AND Date = '" & sDate & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		GetGrpSlotHH = Trim(rsd("OpenFlag" & CStr(CLng(sHH))))
	Else
		GetGrpSlotHH = "Y"
	End If

	rsd.Close

End Function

'��Ə󋵂̎擾
Sub GetOpeStatus(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	iCurTime	[ O ]��ƒ��̎��ԑєԍ��i8�`16,31,99�j
'	iNextTime	[ O ]����Ƃ̎��ԑєԍ��i8�`16,31,99�j
'	iNextStat	[ O ]����Ƃ̃X�e�[�^�X�i0�F�R�O�����O�@1�F�R�O���ȓ��@2�F�T���ȓ��j
'	iOpenSlot	[ O ]�J���g
'	sEndTime	[ O ]�[�ϏI���\��iHHMM�j
	Dim iNextApp

	Call GetOpeStatusDtl(conn, rsd, _
						iCurTime, iNextTime, iNextStat, _
						iOpenSlot, sEndTime, iNextApp)
End Sub

'��Ə󋵏ڍׂ̎擾
Sub GetOpeStatusDtl(conn, rsd, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime, iNextApp)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	iCurTime	[ O ]��ƒ��̎��ԑєԍ��i8�`16,31,99�j
'	iNextTime	[ O ]����Ƃ̎��ԑєԍ��i8�`16,31,99�j
'	iNextStat	[ O ]����Ƃ̃X�e�[�^�X�i0�F�R�O�����O�@1�F�R�O���ȓ��@2�F�T���ȓ��j
'	iOpenSlot	[ O ]�J���g
'	sEndTime	[ O ]�[�ϏI���\��iHHMM�j
'	iNextApp	[ O ]����Ƃ̐\�����i�T���O�m�莞�̒l�j

	Dim sql

	'��Ə�Ԃc�a�Ǎ�
	sql = "SELECT * FROM sOpeStatus"
	rsd.Open sql, conn, 0, 1, 1

	iCurTime = 0
	iNextTime = 0
	If Not rsd.EOF Then		'���R�[�h������ꍇ
		If IsNumeric(rsd("CurTime")) Then		'���l�̏ꍇ
			iCurTime = CLng(rsd("CurTime"))
		Else
			If Trim(rsd("CurTime")) = "B" Then
				iCurTime = 31					'��ƒ��̎��ԑєԍ��i8�`16,31,99�j
			Else
				iCurTime = 99					'��ƒ��̎��ԑєԍ��i8�`16,31,99�j
			End If
		End If
		If IsNumeric(rsd("NextTime")) Then		'���l�̏ꍇ
			iNextTime = CLng(rsd("NextTime"))
		Else
			If Trim(rsd("NextTime")) = "B" Then
				iNextTime = 31					'����Ƃ̎��ԑєԍ��i8�`16,31,99�j
			Else
				iNextTime = 99					'����Ƃ̎��ԑєԍ��i8�`16,31,99�j
			End If
		End If
		iNextStat = CLng(rsd("NextStatus"))		'����Ƃ̃X�e�[�^�X
		iOpenSlot = rsd("OpenSlot")				'�J���g
		sEndTime  = Trim(rsd("WorkEndTime"))	'�[�ϏI���\��iHHMM�j
		iNextApp  = rsd("NextApp")				'����Ƃ̐\�����i�T���O�m�莞�̒l�j
	End If

	rsd.Close

End Sub

'���ԑєԍ����玞�ԑыL�����擾
Function TermNumToSym(iTernNum)
'	iTernNum	[ I ]���ԑєԍ��i8�`16,30,31,32,99�j
'	�߂�l		[ O ]���ԑыL���i08�`16,A,B,D,Z�j

	Select Case iTernNum
   		Case 30		'�ߌ�
			TermNumToSym = "A"
    	Case 31		'�[��
			TermNumToSym = "B"
    	Case 32		'�폜
			TermNumToSym = "D"
    	Case 99		'�Ȃ�
			TermNumToSym = "Z"
		Case Else
			TermNumToSym = Right("0" & CStr(iTernNum), 2)
	End Select 

End Function

'���ԑыL�����玞�ԑєԍ����擾
Function TermSymToNum(sTermSym)
'	sTermSym	[ I ]���ԑыL���i08�`16,A,B,D,Z�j
'	�߂�l		[ O ]���ԑєԍ��i8�`16,30,31,32,99�j
	If IsNumeric(sTermSym) Then
		TermSymToNum = CLng(sTermSym)
	Else
		Select Case sTermSym
   			Case "A"	'�ߌ�
				TermSymToNum = 30
    		Case "B"	'�[��
				TermSymToNum = 31
    		Case "D"	'�폜
				TermSymToNum = 32
    		Case "Z"	'�Ȃ�
				TermSymToNum = 99
		End Select 
	End If
End Function

'�J���g�̌��Z
Function SubmitOpenSlot(conn, rsd)
	'conn		[ I ]DB�ڑ�Object
	'rsd		[ I ]���R�[�h�Z�b�gObject
	'�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sOpeStatus"
	rsd.Open sql, conn, 0, 2, 1

	If Not rsd.EOF Then
		If rsd("OpenSlot") > 0 Then
			rsd("OpenSlot") = rsd("OpenSlot") - 1
			rsd.Update
		End If
	End If

	rsd.Close

End Function

'�\�����̎擾�i�w����A�w��O���[�v�j
'���[�U�������[�U�e�[�u�����擾
Sub GetAppInfo(conn, rsd, sGrpID, WorkDate)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	sGrpID		[ I ]�O���[�v�R�[�h
'	WorkDate	[ I ]��Ɠ�(YYYYMMDD�AWAIT�F�������҂�)

	Dim sql
	Dim sDateNow

	If WorkDate <> "WAIT" Then
		sql = _
		"SELECT sAppliInfo.*, sMUser.Name FROM sAppliInfo, sMUser" & _
		" WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'" & _
		  " AND sAppliInfo.WorkDate = '" & WorkDate & "'" & _
		  " AND RTRIM(sAppliInfo.UserID) = sMUser.UserID" & _
		" Order By sAppliInfo.DelFlag, sAppliInfo.Term, " & _
				 " sAppliInfo.OpeOrder, sAppliInfo.HopeTerm, sAppliInfo.AppDate"
	Else
		'���ݓ��t�擾
		sDateNow = GetYMDStr(Date())

		'�������҂�����
		''(���o�łr�x�ɂ��芮���܂��͋�o���łr�x�ɂ��芮��)�@���邢��
		''�����łr�x�ɂ���폜���L�����Z�����ߋ��̔�����
		sql = _
		"SELECT sAppliInfo.*, sMUser.Name as UName, sTerminal.Name as TName FROM sAppliInfo, sMUser, sTerminal" & _
		" WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'" & _
		  " AND sAppliInfo.Place = 'SY'" & _
		  " AND RTRIM(sAppliInfo.UserID) = RTRIM(sMUser.UserID)" & _
		  " AND RTRIM(sAppliInfo.Terminal) = RTRIM(sTerminal.Terminal)" & _
		  " AND ((sAppliInfo.RecDel = 'D' AND sAppliInfo.Status = '04' or sAppliInfo.RecDel = 'V' AND sAppliInfo.Status = '04' )" & _
		" OR (sAppliInfo.RecDel = 'R' AND" & " (sAppliInfo.Status = '03'" & _
		" OR sAppliInfo.DelFlag = 'Y'" & _
		" OR sAppliInfo.WorkDate < '" & sDateNow & "')))" & _
		" Order By sAppliInfo.WorkDate, sAppliInfo.Term, sAppliInfo.OpeOrder"
	End If
	rsd.Open sql, conn, 0, 1, 1

End Sub

'�\�����̎擾�i�w����A�S�O���[�v�j
Function GetAppInfoAll(conn, rsd, WorkDate)
	'conn		[ I ]DB�ڑ�Object
	'rsd		[I/O]���R�[�h�Z�b�gObject
	'WorkDate	[ I ]��Ɠ�
	'�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
		  " Order By DelFlag, Term, OpeOrder"
	rsd.Open sql, conn, 0, 1, 1

End Function

'�\�����̎擾�i�w����A�w��O���[�v�j
Sub GetAppInfoDG(conn, rsd, sGrpID, WorkDate)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	sGrpID		[ I ]�O���[�v�R�[�h
'	WorkDate	[ I ]��Ɠ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND RTRIM(GroupID) = '" & sGrpID & "'" & _
			" AND DelFlag <> 'Y'" & _
			" AND Status = '02'" & _
		  " Order By DelFlag, Term, OpeOrder"
	rsd.Open sql, conn, 0, 1, 1

End Sub

'�\�����̎擾�i�w����A�w�莞�ԑсA�w��O���[�v�j
Function GetAppInfoTerm(conn, rsd, WorkDate, sHH, sGrpID)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	WorkDate	[ I ]��Ɠ�
'	sHH			[ I ]���ԑсiHH�j
'	sGrpID		[ I ]�O���[�vID
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND Term = '" & sHH & "'" & _
		    " AND DelFlag  <> 'Y'"
	If sGrpID <> "" Then
		sql = sql & " AND RTRIM(GroupID) = '" & sGrpID & "'"
	End If
	rsd.Open sql, conn, 0, 1, 1

End Function

'�\�����̎擾�i�w����A�w�莞�ԑсj
Function GetAppInfoTermAll(conn, rsd, WorkDate, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	WorkDate	[ I ]��Ɠ�
'	sHH			[ I ]���ԑсiHH�j
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND Term = '" & sHH & "'" & _
		    " AND DelFlag  <> 'Y'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'�\�����̎擾�i�w���Ɣԍ��j
Function GetAppInfoOpeNo(conn, rsd, iOpeNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	iOpeNo		[ I ]��Ɣԍ�
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iOpeNo)
	rsd.Open sql, conn, 0, 1, 1

End Function

'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
Function GetAppInfoOpeNoUpd(conn, rsd, iOpeNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	iOpeNo		[ I ]��Ɣԍ�
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iOpeNo)
	rsd.Open sql, conn, 0, 2, 1

End Function

'�\�����̎擾�i�w����A�w��V���[�V�j
Function GetAppInfoChassis(conn, rsd, WorkDate, sChassisID)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	WorkDate	[ I ]��Ɠ�
'	sChassisID	[ I ]�V���[�VID
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE WorkDate = '" & WorkDate & "'" & _
			" AND ChassisId = '" & sChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'�V���g����Ɛ��̎擾�i�w����A�w�莞�ԑсj
Function GetAppWorkCnt(conn, rsd, WorkDate, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	WorkDate	[ I ]��Ɠ�
'	sHH			[ I ]���ԑсiHH�j
'	�߂�l		[ O ]�V���g����Ɛ��i�������j

	Dim iCnt

	iCnt = 0
	'�\�����擾�i�w����A�w�莞�ԑсj�폜�ȊO
	Call GetAppInfoTermAll(conn, rsd, WorkDate, sHH)
	Do Until rsd.EOF
		If rsd("DualOpeNo") = 0 Then
			iCnt = iCnt + 2
		Else
			iCnt = iCnt + 1
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	GetAppWorkCnt = iCnt / 2
End Function

'�V�K��Ə��ʂ̎擾�i�w����A�w�莞�ԑсj
Function GetNewOpeOrder(conn, rsd, WorkDate, sHH, sRD)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	WorkDate	[ I ]��Ɠ�
'	sHH			[ I ]���ԑсiHH�j
'   sRD			[ I ]��ƃ^�C�v�iR/D/V�j'��o���͔��o�Ƃ���
'	�߂�l		[ O ]�V�K��Ə���

	Dim iCnt

	If sHH = "12" Or sHH = "A" Or sHH = "B" Then
		GetNewOpeOrder = 0
		Exit Function
	End If

	iCnt = 0
	'�\�����擾�i�w����A�w�莞�ԑсj�폜�ȊO
	Call GetAppInfoTermAll(conn, rsd, WorkDate, sHH)
	Do Until rsd.EOF
		If rsd("OpeOrder") > iCnt And _
		   rsd("OpeOrder") < 999 Then
			iCnt = rsd("OpeOrder")
		End If
		rsd.MoveNext
	Loop
	rsd.Close

	iCnt = iCnt + 1
	If Left(sRD, 1) = "R" Then	'�����͊
		If (iCnt Mod 2) = 0 Then
			iCnt = iCnt + 1
		End If
	Else						'���o�E��o���͋���
		If (iCnt Mod 2) = 1 Then
			iCnt = iCnt + 1
		End If
	End If

	GetNewOpeOrder = iCnt
End Function

'�\���̗L���`�F�b�N�i�w��R���e�i�j
''�w��R���e�i���L���Ȑ\���Ƃ��ēo�^�����`�F�b�N����
Function ChkAppInfoExt(conn, rsd, sContNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sContNo		[ I ]�R���e�i�ԍ�
'	�߂�l		[ O ]True:���� False:�Ȃ��i�폜�j

	Dim sql
	Dim sDateNow

	'���ݓ��t�擾
	sDateNow = GetYMDStr(Date())

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE RTRIM(ContNo) = '" & sContNo & "'" & _
		    " AND Status = '02'" & _
		    " AND DelFlag <> 'Y'" & _
		    " AND WorkDate >= '" & sDateNow & "'"
	rsd.Open sql, conn, 0, 1, 1

	ChkAppInfoExt = not rsd.EOF
	rsd.Close
End Function

'�\���̗L���`�F�b�N�i�w���Ɣԍ��j
Function ChkAppInfoOpeNo(conn, rsd, iOpeNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	iOpeNo		[ I ]��Ɣԍ�
'	�߂�l		[ O ]True:���� False:�Ȃ�

	Dim sql

	sql = "SELECT * FROM sAppliInfo" & _
		  " WHERE OpeNo = " & CStr(iOpeNo)
	rsd.Open sql, conn, 0, 1, 1

	ChkAppInfoOpeNo = not rsd.EOF
	rsd.Close
End Function

'�c�Ɠ��̎擾
Function GetBusinessDays(conn, rsd, sDateNow, iDaysCnt, sDays())
	'conn		[ I ]DB�ڑ�Object
	'rsd		[ I ]���R�[�h�Z�b�gObject
	'sDateNow	[ I ]�{��
	'iDaysCnt	[ O ]�c�Ɠ���
	'sDays		[ O ]�c�Ɠ�
	'�߂�l		[ O ]�Ȃ�

	Dim sql, sDate, iCnt

	'�J�����_�[�擾
	sql = "SELECT Date FROM sCalendar WHERE HolydayFlag != 'Y' order by Date"
	rsd.Open sql, conn, 0, 1, 1

	iDaysCnt = 0
	'�O�c�Ɠ��擾
	Do Until rsd.EOF
		sDate = Trim(rsd("Date"))
		If sDate >= sDateNow Then
			Exit Do
		End if
		sDays(iDaysCnt) = sDate
		rsd.MoveNext
	Loop
	iDaysCnt = iDaysCnt + 1

	'�������c�Ɠ��Ȃ�ǉ�
	If sDate = sDateNow Then
		sDays(iDaysCnt) = sDate
		iDaysCnt = iDaysCnt + 1
	End if

	'�U����܂ł̉c�Ɠ��擾
	iCnt = 0
	rsd.MoveFirst
	Do Until rsd.EOF
		sDate = Trim(rsd("Date"))
		If sDate > sDateNow Then
			sDays(iDaysCnt) = sDate
			iDaysCnt = iDaysCnt + 1
			iCnt = iCnt + 1
		End if
		If iCnt = 6 Then
			Exit Do
		End if
		rsd.MoveNext
	Loop
	rsd.Close

End Function

'�󂫃X���b�g�̎擾
Function GetLackChassis(conn, rsd, sGrpID, _
			iEmptySlot, iEmptyChassis20, iEmptyChassis40)
	'conn				[ I ]DB�ڑ�Object
	'rsd				[ I ]���R�[�h�Z�b�gObject
	'sGrpID				[ I ]�O���[�vID
	'iEmptySlot			[ O ]�󂫃X���b�g��
	'iEmptyChassis20	[ O ]�󂫃V���[�V��
	'iEmptyChassis40	[ O ]�󂫃V���[�V��

	Dim sql

	sql = "SELECT * FROM sLackChassis" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		iEmptySlot      = rsd("EmptySlot")
		iEmptyChassis20 = rsd("EmptyChassis20")
		iEmptyChassis40 = rsd("EmptyChassis40")
	Else	'�󂫂Ȃ��̏ꍇ�̓��R�[�h���Ȃ�
		iEmptySlot      = 0
		iEmptyChassis20 = 0
		iEmptyChassis40 = 0
	End If
	rsd.Close

End Function

'�V���g�����ɂ��O���[�v���蓖�Ęg�̎擾
Function GetShuttleUnit(conn, rsd, sYMD, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sYMD		[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	�߂�l		[ O ]�O���[�v���蓖�Ęg

	Dim sql
	Dim iWk, iQty
	
	iWk = Weekday(cDate(ChgYMDStr(sYMD)))
	iQty = 0
	sql = "SELECT * FROM sShuttleUnit" & _
		  " WHERE Weekday = " & CStr(iWk)
	rsd.Open sql, conn, 0, 1, 1

	iQty = 0
	If Not rsd.EOF Then
		If sHH < "12" Then
			iQty = rsd("reservAM")
		Else
			iQty = rsd("reservPM")
		End If
	End If
	rsd.Close

	GetShuttleUnit = iQty

End Function

'�O���[�v���̎擾
Function GetGrpCnt(conn, rsd)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	�߂�l		[ O ]�O���[�v��

	Dim sql
	Dim iWk

	sql = "SELECT GroupID FROM sMGroup"
	rsd.Open sql, conn, 0, 1, 1

	iWk = 0
	Do Until rsd.EOF
		iWk = iWk + 1
		rsd.MoveNext
	Loop

	rsd.Close

	GetGrpCnt = iWk

End Function

'���ԑѕ��̎擾
Function GetTimeSlotSize(conn, rsd, sHH, iTime)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	iTime		[ O ]���ԑѕ��i���j
'	�߂�l		[ O ]�Ȃ�

	Dim sWk

	'���ԑѕ����v�Z
	iTime = 60
	sWk = GetEnv(conn, rsd, "ShtStart")
	If sHH = Left(sWk, 2) Then
		iTime = 60 - CLng(Right(sWk, 2))
	End If
	sWk = GetEnv(conn, rsd, "ShtEnd")
	If sHH = Left(sWk, 2) Then
		iTime = CLng(Right(sWk, 2))
	End If
'''2001/08/20 debug
'''Response.Write "���ԑѕ�(" & iTime & ")"
'''
End Function

'���o���R���e�i�̎擾
Function GetOutCont(conn, rsd, sContNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	sContNo		[ I ]�R���e�i�ԍ�
'	�߂�l		[ O ]�Ȃ�

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sql = "SELECT * FROM sOutContainer" & sNo & _
		  " WHERE RTRIM(ContNo) = '" & sContNo & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'���o���}�j�t�F�X�g�R���e�i�̎擾
Function GetOutBLCont(conn, rsd, sBLNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	sBLNo		[ I ]BL�ԍ�
'	�߂�l		[ O ]�Ȃ�

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sql = "SELECT * FROM sOutBLCont" & sNo & _
		  " WHERE RTRIM(BLNo) = '" & sBLNo & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'���o���}�j�t�F�X�g�R���e�i�ƃR���e�i�̎擾(�����b�x�Ή�)01/10/01
Function GetOutBLContCont(conn, rsd, sBLNo, sTerm_CD)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	sBLNo		[ I ]BL�ԍ�
'	sTerm_CD	[ I ]�^�[�~�i���R�[�h
'	�߂�l		[ O ]�Ȃ�

	Dim sql, sNo, sTbl1, sTbl2

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sTbl1 = "sOutContainer" & sNo
	sTbl2 = "sOutBLCont" & sNo

	sql = "SELECT " & sTbl1 & ".*" & _
		  " FROM " & sTbl1 & "," & sTbl2 & _
		  " WHERE RTRIM(" & sTbl2 & ".BLNo) = '" & sBLNo & "'" & _
		    " AND RTRIM(" & sTbl2 & ".ContNo) = RTRIM(" & sTbl1 & ".ContNo)" & _
		    " AND RTRIM(" & sTbl1 & ".Terminal) = '" & sTerm_CD & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'�u�b�L���O�̗L������(02/02/09)
Function ChkBooking(conn, rsd, sBookNo, sOpeRec, sVslCode, sVoyage)
'	conn		[ I ]DB�ڑ�Object
'	rsd		[ I ]���R�[�h�Z�b�gObject
'	sBookNo		[ I ]�u�b�L���O�ԍ�
'	sOpeRec		[ O ]�����I�y���[�^
'	sVslCode	[ O ]�����D��		'02/02/09 ADD
'	sVoyage		[ O ]�������q		'02/02/09 ADD
'	�߂�l		[ O ]�u�b�L���O�̗L��

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	sql = "SELECT * FROM sBooking" & sNo & _
		  " WHERE RTRIM(BookNo) = '" & sBookNo & "'"
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
		sOpeRec = Trim(rsd("Operator"))
'02/02/09 ADD
		sVslCode = Trim(rsd("VslCode"))	
		sVoyage = Trim(rsd("Voyage"))
'02/02/09 ADD
	End If
	ChkBooking = Not rsd.EOF
	rsd.Close

End Function

'�V���[�V�̎擾
Function GetChassis(conn, rsd, sChassisID)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[I/O]���R�[�h�Z�b�gObject
'	sChassisID	[ I ]�V���[�VID
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	sql = "SELECT * FROM sChassis" & _
		  " WHERE RTRIM(ChassisId) = '" & sChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1

End Function

'�O���[�v�̊��蓖�Ęg���擾
Function GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sGrpID		[ I ]�O���[�v�R�[�h
'	sYMD		[ I ]��Ɠ��iYYYYMMDD�j
'	sHH			[ I ]���ԑсiHH�j
'	�߂�l		[ O ]�O���[�v�̊��蓖�Ęg
	Dim iCnt, iTime

	'�O���[�v���蓖�Ęg�̎擾
	iCnt = GetShuttleUnit(conn, rsd, sYMD, sHH)
'''2001/08/20 debug
'''Response.Write "�O���[�v���蓖�Ęg(" & iCnt & ")"
'''

	'���ԑѕ��̎擾
	Call GetTimeSlotSize(conn, rsd, sHH, iTime)

	GetGrpAssign = iCnt * iTime / 60

'''2001/08/20 debug
'''Response.Write "���蓖�Ęg(" & GetGrpAssign & ")"
'''

End Function

'��Ɣԍ��̔���
Function GetNewOpeNo(conn, rsd, sYMD, sOpeNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sYMD		[ I ]�V���[�VID
'	sOpeNo		[ O ]��Ɣԍ�
'	�߂�l		[ O ]�Ȃ�

	Dim bOk

	bOk = False
	Do Until bOk
		'��Ɣԍ��̔���
		Call GetNewOpeNo2(conn, rsd, sYMD, sOpeNo)

		'�\���̗L���`�F�b�N�i�w���Ɣԍ��j
		If Not ChkAppInfoOpeNo(conn, rsd, CLng(sOpeNo)) Then
			bOk = True
		End If
	Loop

End Function

'��Ɣԍ��̔���
Function GetNewOpeNo2(conn, rsd, sYMD, sOpeNo)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sYMD		[ I ]�V���[�VID
'	sOpeNo		[ O ]��Ɣԍ�
'	�߂�l		[ O ]�Ȃ�

	Dim sql, iVal

	sql = "SELECT * FROM sMAutoNumber" & _
		  " WHERE KindCode = '01'" & _
		    " AND KeyCode = '" & Right(sYMD, 2) & "'"
	rsd.Open sql, conn, 0, 2, 1

	If rsd.EOF Then
		'�V�K�̏ꍇ
		rsd.AddNew
		rsd("KindCode") = "01"
		rsd("KeyCode") = Right(sYMD, 2)
		rsd("InitVal") = 1
		rsd("CurrentVal") = 1
		rsd("MaxVal") = 999
		rsd("TrgDate") = sYMD
	End If

	'���t���ς�����ꍇ
	If rsd("TrgDate") <> sYMD Then
		rsd("CurrentVal") = rsd("InitVal")
		rsd("TrgDate") = sYMD
	End If

	'��Ɣԍ��擾
	iVal = rsd("CurrentVal")

	If iVal = rsd("MaxVal") Then
		'�ő�l�̏ꍇ
		rsd("CurrentVal") = rsd("InitVal")
	Else
		rsd("CurrentVal") = iVal + 1
	End If
	rsd("UpdtTime") = Now()
	rsd("UpdtPgCd") = "SYWB012"
	rsd("UpdtTmnl") = "ME"

	rsd.Update
	rsd.Close

	sOpeNo = Right(sYMD, 2) & Right("000" & CStr(iVal), 3)
	GetNewOpeNo2 = sOpeNo
End Function

'�j�����擾����
Function sWeekday(iWk)
'	sYMD		[ I ]���l
'	�߂�l		[ O ]�j��

	select case	iWk
        Case 1
            sWeekday =  "��"
        Case 2
            sWeekday =  "��"
        Case 3
            sWeekday =  "��"
        Case 4
            sWeekday =  "��"
        Case 5
            sWeekday =  "��"
        Case 6
            sWeekday =  "��"
        Case 7
            sWeekday =  "�y"
	end select
End Function

'�\����A�\�񎞊ԑт��L�����`�F�b�N
Function CheckAppWorkDate(conn, rsd, sYMD, sHH, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd				[ I ]���R�[�h�Z�b�gObject
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH				[ I ]�\�񎞊ԑсiHH�j
'	sErrMsg			[ O ]�G���[���b�Z�[�W

	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime, sWk

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
		sWk = Right("0" & CStr(iNextTime), 2)
		If sHH < sWk or (sHH = sWk and iNextStat = 2) Then
			sErrMsg = "�w�莞�ԑт̗\��͒��ߐ؂��Ă��܂�"
			Exit Function
		End If
	End If
End Function

'�Œ茅���ł̕\��
Function GetNumStr(vNum, iK)
'	vNum			[ I ]���l�i���lor������j
'	iK				[ I ]����
	GetNumStr = Right(String(iK, "0") & CStr(vNum), iK)
End Function

'�^�[�~�i���̎擾
Function GetTerminal(conn, rsd, sTerminal())
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sTerminal	[ O ]�^�[�~�i������(����3�����O)
'	�߂�l		[ O ]�Ȃ�

	Dim sql
	Dim i

	'�^�[�~�i���擾
	sql = "SELECT * FROM sTerminal WHERE Umu <> '1' "
	sql = sql & "  Order By Terminal"
	rsd.Open sql, conn, 0, 1, 1

	i = 0
	Do Until rsd.EOF
		sTerminal(i) = rsd("Terminal") & Trim(rsd("Name"))
		i = i + 1
		rsd.MoveNext
	Loop
	rsd.Close


End Function

'�R���e�i�^�C�v���̂̎擾
Function GetContType(conn, rsd, sContType())
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sContType	[ O ]�R���e�i�^�C�v����(����2�����O)
'	�߂�l		[ O ]�Ȃ�

	Dim sql
	Dim i

	'�R���e�i�^�C�v�擾
	sql = "SELECT * FROM sContType" 
	rsd.Open sql, conn, 0, 1, 1

	i = 0
	Do Until rsd.EOF
		sContType(i) = rsd("ContType") & Trim(rsd("Name"))
		i = i + 1
		rsd.MoveNext
	Loop
	rsd.Close


End Function

'�R���e�i�ގ��擾�̎擾
Function GetMaterial(conn, rsd, sMaterial())
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sMaterial	[ O ]�R���e�i�ގ�(����3�����O)
'	�߂�l		[ O ]�Ȃ�

	Dim sql
	Dim i

	'�R���e�i�ގ��擾
	sql = "SELECT * FROM sMaterial" 
	rsd.Open sql, conn, 0, 1, 1

	i = 0
	Do Until rsd.EOF
		sMaterial(i) = rsd("Material") & Trim(rsd("Name"))
		i = i + 1
		rsd.MoveNext
	Loop
	rsd.Close

End Function


'�^�[�~�i�����̓��̎擾
Function GetTerminal2(conn, rsd, sTerminal, sType, sName)
'	conn		[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sTerminal	[ I ]�^�[�~�i���R�[�h
'	sType		[ O ]�^�C�v
'	sName		[ O ]�^�[�~�i������
'	�߂�l		[ O ]�Ȃ�

	Dim sql

	'�^�[�~�i���擾
	sql = "SELECT Terminal, Type, Name FROM sTerminal" & _
		  " WHERE RTRIM(Terminal) = '" & sTerminal & "'"
	rsd.Open sql, conn, 0, 1, 1

	sType  = Trim(rsd("Type"))		
	sTerm_Name  = Trim(rsd("Name"))
	rsd.Close
End Function

%>