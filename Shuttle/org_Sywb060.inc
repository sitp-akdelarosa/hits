<% 

'��V���[�V�r�x�݌ɂ̔���
Function GetVPBooking1(conn, rsd, sVPBook1, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sVPBook1		[ I ]�u�b�L���O�ԍ�
'	sErrMsg			[ O ]�G���[���b�Z�[�W�i�����̏ꍇ�͋󔒁j

	Dim sql, sNo

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	'SELECT���쐬
	sql = "SELECT VslCode, Voyage, LineCode, BookNo, Terminal, ContSize, "
    	sql = sql & " ContType, ContHeight, Material FROM sVPBooking" & sNo
    	sql = sql & " WHERE RTRIM(BookNo) = '" & sVPBook1 & "'" 
	sql = sql & " GROUP BY VslCode, Voyage, LineCode, BookNo, Terminal, "
	sql = sql & " ContSize, ContType, ContHeight, Material ; "
	rsd.Open sql, conn, 0, 1, 1

	IF rsd.EOF Then
		sErrMsg = "�Y������u�b�L���O��񂪂���܂���"
		rsd.Close
	End If

End Function

'�u�b�L���O�\��`�F�b�N�����P�i��Ɠ��̃`�F�b�N�j
Function VPBookCheck1(conn, rsd, _
			sUsrID, sGrpID,	sYMD, sHH, _
			sVPBook, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	sVPBook			[ I ]�����u�b�L���O�ԍ�
'	sErrMsg			[ O ]�G���[���b�Z�[�W�i�����̏ꍇ�͋󔒁j

	Dim sWk
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

	End if
End Function

'�u�b�L���O�\��X�V����
Function UpdateApp_VPBook(conn, rsd, sUsrID, sGrpID, _
			sYMD, sHH, sVSL, sVOY, sLINE, sVPBook, _
            sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
            sErrMsg, sOpeNoVan)

'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	sVSL			[ I ]�D��
'	sVOY			[ I ]���q
'	sLINE			[ I ]�q�H
'	sVPBook			[ I ]��o���u�b�L���O�ԍ�
'	sTERM			[ I ]�^�[�~�i���R�[�h
'	sSIZE			[ I ]�R���e�i�T�C�Y
'	sTYPE			[ I ]�^�C�v
'	sHIGHT			[ I ]����
'	sMATERIAL		[ I ]�ގ�
'	sDeliverTo		[ I ]�R���e�i���o��
'	sErrMsg			[ O ]�G���[���b�Z�[�W�i�����̏ꍇ�͋󔒁j
'	sOpeNoVan		[ O ]��o���\��ԍ�

	Dim i, sWk, iAssignCnt, bOk, iDualOpeNo, iOpeOrder, iCnt
	Dim sDateNow, iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime
	Dim sVslCode, sVoyage, sOpenFlag

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
	End if

	'���ԑюw��̏ꍇ�ɂ̓O���[�v�̊��蓖�Ęg���擾
	iAssignCnt = -1
	sOpenFlag = ""
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'�O���[�v�̊��蓖�Ęg���擾
		iAssignCnt = GetGrpAssign(conn, rsd, sGrpID, sYMD, sHH)

		'�O���[�v���ԑя��̎擾
		sOpenFlag = GetGrpSlotHH(conn, rsd, sGrpID, sYMD, sHH)

	End If

	'�\��`�F�b�N
	bOk = AppUpdCheck_VPBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
                             iCnt, sOpeVP, sErrMsg) 

	If bOk Then
		'�\�񏈗�
		Call UpdateApp_VPBook2(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
			                 iCnt, sOpeVP, sErrMsg, sOpeNoVan) 
	End If

End Function

'�\��`�F�b�N
Function AppUpdCheck_VPBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
                             iCnt, sOpeVP, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	sVSL			[ I ]�D��
'	sVOY			[ I ]���q
'	sLINE			[ I ]�q�H
'	sVPBook			[ I ]��o���u�b�L���O�ԍ�
'	sTERM			[ I ]�^�[�~�i���R�[�h
'	sSIZE			[ I ]�R���e�i�T�C�Y
'	sTYPE			[ I ]�^�C�v
'	sHIGHT			[ I ]����
'	sMATERIAL		[ I ]�ގ�
'	iAssignCnt 		[ I ]�O���[�v������
'	iOpenSlot 		[ I ]�J���g��
'	iDualOpeNo		[ O ]�f���A����Ɣԍ�
'	iOpeOrder		[ O ]��Ə���
'	iCnt			[ O ]�\��g��
'	sOpeVP			[ O ]�I�y���[�^
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheck_VPBook = False

	'���o�\�u�b�L���O�f�[�^�`�F�b�N
	If Not AppUpdCheckBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, sTERM, _
                             sSIZE, sTYPE, sHIGHT, sMATERIAL, sOpeVP, sErrMsg) Then
		Exit Function
	End If

	'���ԑюw��̏ꍇ�͘g���`�F�b�N
	If sHH <> "12" and sHH <> "A" and sHH <> "B" Then
		'�\��g�ɂ��`�F�b�N
		If Not AppUpdCheckTerm_VP(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, sTERM, _
						     sSIZE, sTYPE, sHIGHT, sMATERIAL, iAssignCnt, iOpenSlot, _
						     iDualOpeNo, iOpeOrder, iCnt, sErrMsg) Then

			Exit Function
		End If
	Else
		iDualOpeNo = 0		'�f���A����Ɣԍ�
		iOpeOrder = 0		'��Ə���
		iCnt = 0			'�\��g��
	End If

	AppUpdCheck_VPBook = True
End Function

'���o�\�u�b�L���O�f�[�^�`�F�b�N
Function AppUpdCheckBook(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			             sVSL, sVOY, sLINE, sVPBook, sTERM, _
                         sSIZE, sTYPE, sHIGHT, sMATERIAL, sOpeVP, sErrMsg) 
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	sVSL			[ I ]�D��
'	sVOY			[ I ]���q
'	sLINE			[ I ]�q�H
'	sVPBook			[ I ]��o���u�b�L���O�ԍ�
'	sTERM			[ I ]�^�[�~�i���R�[�h
'	sSIZE			[ I ]�R���e�i�T�C�Y
'	sTYPE			[ I ]�^�C�v
'	sHIGHT			[ I ]����
'	sMATERIAL		[ I ]�ގ�
'	sOpeVP          	[ O ]�I�y���[�^
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheckBook = False

	Dim sql, sNo
	Dim iRestQtyBook, iCount '�\��c��

	rsd.Open "sUseDB", conn, 0, 1, 2
	If not rsd.EOF Then
		sNo = rsd("EnableDB")
	End If
	rsd.Close

	'SELECT���쐬
	sql = "SELECT RestQtyBook, Operator FROM sVPBooking" & sNo
    	sql = sql & " WHERE RTRIM(VslCode) = '" & sVSL & "'" 
    	sql = sql & " AND   RTRIM(Voyage) = '" & sVOY & "'" 
    	sql = sql & " AND   RTRIM(LineCode) = '" & sLINE & "'" 
    	sql = sql & " AND   RTRIM(BookNo) = '" & sVPBook & "'" 
    	sql = sql & " AND   RTRIM(Terminal) = '" & sTERM & "'" 
    	sql = sql & " AND   RTRIM(ContSize) = '" & sSIZE & "'" 
    	sql = sql & " AND   RTRIM(ContType) = '" & sTYPE & "'" 
    	sql = sql & " AND   RTRIM(ContHeight) = '" & sHIGHT & "'" 
    	sql = sql & " AND   RTRIM(Material) = '" & sMATERIAL & "'" 
	rsd.Open sql, conn, 0, 1, 1

	If not rsd.EOF Then
		iRestQtyBook = rsd("RestQtyBook")
		sOpeVP = Trim(rsd("Operator"))
	Else
		sErrMsg = "�Y������u�b�L���O��񂪂���܂���B"
	End If

	rsd.Close

	If sErrMsg <> "" Then
		Exit Function
	End If

	'SELECT���쐬(�\��{���`�F�b�N)
	sql = "SELECT count(OpeNo) AS CountApp FROM sAppliInfo" 
    	sql = sql & " WHERE RTRIM(VslCode) = '" & sVSL & "'" 
    	sql = sql & " AND   RTRIM(Voyage) = '" & sVOY & "'" 
    	sql = sql & " AND   RTRIM(LineCode) = '" & sLINE & "'" 
    	sql = sql & " AND   RTRIM(VPBookNo) = '" & sVPBook & "'" 
    	sql = sql & " AND   RTRIM(Terminal) = '" & sTERM & "'" 
    	sql = sql & " AND   RTRIM(ContSize) = '" & sSIZE & "'" 
    	sql = sql & " AND   RTRIM(ContType) = '" & sTYPE & "'" 
    	sql = sql & " AND   RTRIM(ContHeight) = '" & sHIGHT & "'" 
    	sql = sql & " AND   RTRIM(Material) = '" & sMATERIAL & "'" 
    	sql = sql & " AND   DelFlag <> 'Y' "
    	sql = sql & " AND   Status = '02' "
	rsd.Open sql, conn, 0, 1, 1

	iCount = rsd("CountApp")
	rsd.Close

	If iRestQtyBook <= iCount Then
		sErrMsg = "�{���s���ɂ���o���s�b�N�ł��܂���"
		Exit Function
	End If
	AppUpdCheckBook = True
End Function

'�\��g�ɂ��`�F�b�N(�o���v�[��)

Function AppUpdCheckTerm_VP(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, sTERM, _
						     sSIZE, sTYPE, sHIGHT, sMATERIAL, iAssignCnt, iOpenSlot, _
						     iDualOpeNo, iOpeOrder, iCnt, sErrMsg)
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	sVSL			[ I ]�D��
'	sVOY			[ I ]���q
'	sLINE			[ I ]�q�H
'	sVPBook			[ I ]��o���u�b�L���O�ԍ�
'	sTERM			[ I ]�^�[�~�i���R�[�h
'	sSIZE			[ I ]�R���e�i�T�C�Y
'	sTYPE			[ I ]�^�C�v
'	sHIGHT			[ I ]����
'	sMATERIAL		[ I ]�ގ�
'	iAssignCnt 		[ I ]�O���[�v������
'	iOpenSlot 		[ I ]�J���g��
'	iDualOpeNo		[ O ]�f���A����Ɣԍ�
'	iOpeOrder		[ O ]��Ə���
'	iCnt			[ O ]�\��g��
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	�߂�l			[ O ]OK/NG

	AppUpdCheckTerm_VP = False

	'�\�����̎擾�i�w����A�w�莞�ԑсA�w��O���[�v�j
	Call GetAppInfoTerm(conn, rsd, sYMD, sHH, sGrpID)
	iCnt = 0			'���݉����񐔁i���o�^�O���[�v�g���j
	iDualOpeNo = 0			'�f���A����Ɣԍ�
	iOpeOrder = 0			'��Ə���
	Do Until rsd.EOF
		'�����Ńf���A���ɂȂ��ꍇ
		'�y�A�ɂȂ蓾����o���`�F�b�N

		If rsd("DualOpeNo") = 0 and _
		   rsd("RecDel") = "R" and _
			rsd("NotDelFlag") <> "Y" and _
		   Trim(rsd("ChassisId")) = "" Then
					'�������O�󂯂̏ꍇ�́A�ΏۂƂȂ��o�����o�͓����^�[�~�i���Ɍ���
					'If (Trim(rsd("Terminal")) = "KA" OR  Trim(rsd("Terminal")) = "HK") OR _
					'   (Trim(rsd("Terminal")) <> "KA" AND Trim(rsd("Terminal")) <> "HK" AND sTERM = Trim(rsd("Terminal"))) Then	
					'	'���p���T�C�Y���}�b�`����ꍇ
					'	If (rsd("ContSize") =  "20" and sSIZE =  "20") or _
					'	   (rsd("ContSize") <> "20" and sSIZE <> "20") Then
					'		iDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
					'		'01/10/30''
					'		'''iOpeOrder = rsd("OpeOrder") - 1
					'		iOpeOrder = rsd("OpeOrder") + 1
					'	End If
					'End If
			'DUAL����^�[�~�i���Ƃ���(02/04/27)
			If sTERM = Trim(rsd("Terminal")) Then	
				'���p���T�C�Y���}�b�`����ꍇ
				If (rsd("ContSize") =  "20" and sSIZE =  "20") or _
				   (rsd("ContSize") <> "20" and sSIZE <> "20") Then
					iDualOpeNo = rsd("OpeNo")	'�f���A����Ɣԍ�
					'01/10/30''
					'''iOpeOrder = rsd("OpeOrder") - 1
					iOpeOrder = rsd("OpeOrder") + 1
				End If
			End If

		End If

'''		If rsd("DualOpeNo") = 0 or _			'01/10/30 bug
'''		   rsd("RecDel") = "D" Then
		If rsd("DualOpeNo") = 0 or _
		   rsd("RecDel") <> "R" Then
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
		sRDType = "DEL"
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, sRDType)
	End If

	AppUpdCheckTerm_VP = True

End Function


'�\��X�V�����Q
Function UpdateApp_VPBook2(conn, rsd, sUsrID, sGrpID, sYMD, sHH, _
			                 sVSL, sVOY, sLINE, sVPBook, _
                             sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
			                 iAssignCnt, iOpenSlot, iDualOpeNo, iOpeOrder, _
			                 iCnt, sOpeVP, sErrMsg, sOpeNoVan) 
'	conn			[ I ]DB�ڑ�Object
'	rsd			[ I ]���R�[�h�Z�b�gObject
'	sUsrID			[ I ]���[�UID
'	sGrpID			[ I ]�O���[�vID
'	sYMD			[ I ]�\����iYYYYMMDD�j
'	sHH			[ I ]�\�񎞊ԑсiHH�j
'	sVSL			[ I ]�D��
'	sVOY			[ I ]���q
'	sLINE			[ I ]�q�H
'	sVPBook			[ I ]��o���u�b�L���O�ԍ�
'	sTERM			[ I ]�^�[�~�i���R�[�h
'	sSIZE			[ I ]�R���e�i�T�C�Y
'	sTYPE			[ I ]�^�C�v
'	sHIGHT			[ I ]����
'	sMATERIAL		[ I ]�ގ�
'	sDeliverTo		[ I ]�R���e�i���o��
'	iAssignCnt 		[ I ]�O���[�v������
'	iOpenSlot 		[ I ]�J���g��
'	iDualOpeNo		[ I ]�f���A����Ɣԍ�
'	iOpeOrder		[ I ]��Ə���
'	iCnt			[ I ]�\��g��
'	sOpeVP			[ I ]�I�y���[�^
'	sErrMsg			[ O ]�G���[���b�Z�[�W
'	sOpeNoVan		[ O ]��o���\��ԍ�

	Dim sql, bTermFlag

	sErrMsg = "�c�a�X�V�G���[���������܂���"

	sOpeNoVan = "0"

	'��Ɣԍ��̔���
	Call GetNewOpeNo(conn, rsd, sYMD, sOpeNoVan)

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
		rsd("DualOpeNo") 	= CLng(sOpeNoVan)		'�f���A����Ɣԍ�
		rsd("SendFlag")    	= "Y"				'�ύX���M�׸
	End If

	rsd.AddNew
	rsd("OpeNo")       = sOpeNoVan				'��Ɣԍ�
	rsd("UpdtTime")    = Now()				'�X�V����
	rsd("UpdtPgCd")    = "SYWB012"				'�X�V��۸��і�
	rsd("UpdtTmnl")    = "WEB"				'�X�V�[����
	rsd("UserID")      = sUsrID				'���[�U���
	rsd("GroupID")     = sGrpID				'��ٰ�ߺ��
	rsd("AppDate")     = Now()				'�\����
	rsd("WorkDate")    = ChgYMDDate(sYMD)			'��Ɠ�
	rsd("RecDel")      = "V"				'��Ǝ��
	rsd("ContNo")      = ""					'���Ŕԍ�
	rsd("BookNo")      = ""					'�ޯ�ݸޔԍ�
	rsd("Status")      = "02"				'���
	rsd("Place")       = "VP"				'�ꏊ(�o���v�[��)
	rsd("ChassisId")   = ""					'����ID
	rsd("DelFlag")     = ""					'�폜�׸
	rsd("CReason")    = ""					'�L�����Z�����R
	rsd("WorkFlag")    = ""					'���ٍ�ƒ��׸
	rsd("SendFlag")    = "Y"				'�ύX���M�׸
	If sSIZE = "20" Then
		rsd("Size20Flag")  = "Y"			'20�t�B�[�g�t���O
	Else
		rsd("Size20Flag")  = ""				'20�t�B�[�g�t���O
	End If
	rsd("VslCode")     = sVSL				'�D�����
	rsd("Voyage")      = sVOY				'���q
	rsd("BLNo")        = ""					'�a�k�ԍ�
	rsd("Term")        = sHH				'���ԑ�
	If Not bTermFlag Then		'���ԑюw��łȂ��ꍇ
		rsd("OpeOrder") = 0				'��Ə���
	ElseIf iDualOpeNo > 0 Then	'���̍�Ƃƃf���A���ɂȂ�ꍇ
		rsd("OpeOrder") = iOpeOrder			'��Ə���
	Else
		rsd("OpeOrder") = iOpeOrder			'��Ə���
	End If
	rsd("HopeTerm")    = sHH				'��]���ԑ�
	If iDualOpeNo > 0 Then	'���̍�Ƃƃf���A���ɂȂ�ꍇ
		rsd("DualOpeNo") = iDualOpeNo		'�f���A����Ɣԍ�
	Else
		rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
	End If
	rsd("ContSize")    = sSIZE				'�R���e�i�T�C�Y
	rsd("NotDelFlag")  = ""					'���o���ڂ��Ȃ�
	rsd("NightFlag")   = ""					'�[�σV���[�V
	rsd("MixSizeFlag") = ""					'20/40���p�V���[�V
	rsd("DeliverTo")   = sDeliverTo			'���o��
	rsd("LockFlag")    = ""					'���b�N�t���O
	rsd("Operator")    = sOpeVP				'�I�y���[�^�R�[�h
	rsd("OpeBookNo")   = ""					'���Ѓu�b�L���O�ԍ�
	rsd("ReservNo")    = 0					'�\��ԍ�

	rsd("RecContSize") = sSIZE				'���Ż���
	rsd("ContHeight")  = sHIGHT				'����ʲ�
	rsd("ContType")    = sTYPE				'��������
	rsd("Material")    = sMATERIAL			'���ōގ�
	rsd("TareWeight")  = 0					'��������ñ����
	rsd("Gweight")     = 0					'�������Ÿ�۽����
	rsd("SoundUnsound")  = ""				'����Sound/Unsound
	rsd("Trucker")  = ""					'�����^���Ǝ�
	rsd("ReceiveFrom")  = ""				'������
	rsd("CustOK")  = ""						'�������Œʊ�/���ʊ�
	rsd("Terminal")  = sTERM				'�Y���^�[�~�i��(VP�\��)�b��C��
	rsd("VPBookNo")      = sVPBook			'���o�ޯ�ݸޔԍ�(VP�\��)�b��C��
	rsd("LineCode")      = sLINE			'�q�H�R�[�h(VP�\��)�b��C��
	rsd("HeadID")        = ""    			'�w�b�h�h�c(ICCT�Ή�) 03/07/31

	rsd.Update
	rsd.Close
	sErrMsg = ""

End Function

'�^�[�~�i�����̂̎擾
Function GetTerminal_Name(conn, rsd, sTerminal)
'	conn		[ I ]DB�ڑ�Object
'	rsd		[ I ]���R�[�h�Z�b�gObject
'	sTerminal	[ I ]�^�[�~�i���R�[�h
'	�߂�l		[ O ]����

	Dim sql
	Dim i

	GetTerminal_Name = ""
	'�^�[�~�i���擾
	sql = "SELECT Name FROM sTerminal WHERE RTRIM(Terminal) = '" & sTerminal & "'" & _
		    " AND Umu <> '1'"
	rsd.Open sql, conn, 0, 1, 1

	GetTerminal_Name = Trim(rsd("Name"))

	rsd.Close

End Function

'�R���e�i�^�C�v���̂̎擾
Function GetContType_Name(conn, rsd, sContType)
'	conn		[ I ]DB�ڑ�Object
'	rsd		[ I ]���R�[�h�Z�b�gObject
'	sContType	[ I ]�R���e�i�^�C�v�R�[�h
'	�߂�l		[ O ]����

	Dim sql
	Dim i

	GetContType_Name = ""

	'�R���e�i�^�C�v�擾
	sql = "SELECT Name FROM sContType WHERE RTRIM(ContType) = '" & sContType & "'"
	rsd.Open sql, conn, 0, 1, 1

	GetContType_Name = Trim(rsd("Name"))
	rsd.Close


End Function

'�R���e�i�ގ��擾�̎擾
Function GetMaterial_Name(conn, rsd, sMaterial)
'	conn		[ I ]DB�ڑ�Object
'	rsd		[ I ]���R�[�h�Z�b�gObject
'	sMaterial	[ I ]�R���e�i�ގ��R�[�h
'	�߂�l		[ O ]����

	Dim sql
	Dim i

	'�R���e�i�ގ��擾
	sql = "SELECT Name FROM sMaterial WHERE RTRIM(Material) = '" & sMaterial & "'"
	rsd.Open sql, conn, 0, 1, 1

	GetMaterial_Name = Trim(rsd("Name"))
	rsd.Close

End Function

'�Ώۃo���v�[�����̌���
Function SetTerminal(sTerminal, sTerm_Tbl())
	'sTerminal	[ I ]�^�[�~�i���R�[�h
	'sTerm_Tbl()	[ I ]�^�[�~�i�����̃e�[�u��
	'�߂�l		[ O ]�Ώۃ^�[�~�i������

	Dim sStr
	Dim i

	SetTerminal = ""
	for i = 0 to 49
		If sTerminal = Left(sTerm_Tbl(i),3) Then
			SetTerminal = Mid(sTerm_Tbl(i), 4)
			Exit for
		End If
	Next

End Function

'�R���e�i�^�C�v���̌���
Function SetContType(sContType, sContType_Tbl())
	'sContType		[ I ]�R���e�i�^�C�v
	'sContType_Tbl()	[ I ]�R���e�i�^�C�v�e�[�u��
	'�߂�l			[ O ]�ΏۃR���e�i�^�C�v����

	Dim sStr
	Dim i

	SetContType = ""
	for i = 0 to 14
		If sContType = Left(sContType_Tbl(i),2) Then
			SetContType = Mid(sContType_Tbl(i), 3)
			Exit for
		End If
	Next

End Function

'�R���e�i�ގ����̌���
Function SetMaterial(sMaterial, sMaterial_Tbl())
	'sMaterial	[ I ]�R���e�i�ގ�
	'sMaterial()	[ I ]�R���e�i�ގ��e�[�u��
	'�߂�l		[ O ]�ΏۃR���e�i�ގ�����

	Dim sStr
	Dim i

	SetMaterial = ""
	for i = 0 to 9
		If sMaterial = Left(sMaterial_Tbl(i),1) Then
			SetMaterial = Mid(sMaterial_Tbl(i), 2)
			Exit for
		End If
	Next

End Function

%>