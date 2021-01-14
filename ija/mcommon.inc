<%
' ���[�UID�̌`���`�F�b�N
Function CheckUserID(sUserID)
	'sUserID		[ I ]���[�UID
	'CheckUserID	[ O ]�G���[���b�Z�[�W

	' �^����Љ�ЃR�[�h(���[�UID�̐擪)�̍ő�l
	Const COMPANY_MAX = "008"

	Dim sCompany, sDriver

	If sUserID = "" Then
		CheckUserID = "���[�U�[ID������"
		Exit Function
	End If

	If (Not IsNumeric(sUserID)) Or (Len(sUserID) <> 6) Then
		CheckUserID = "�s���ȃ��[�U�[ID<br>[�@" & sUserID & "�@]"
		Exit Function
	End If

	sCompany = Left(sUserID, Len(COMPANY_MAX))
	sDriver = Right(sUserID, Len(sUserID) - Len(COMPANY_MAX))
	If CInt(sCompany) = 0 Or CInt(sDriver) = 0 Or CInt(COMPANY_MAX) < CInt(sCompany) Then
		CheckUserID = "�s���ȃ��[�U�[ID<br>[�@" & sUserID & "�@]"
		Exit Function
	End If

	CheckUserID = ""
End function

' �g�ѓd�b�̎�ʂ��擾
Function GetPhoneType()
	'GetPhoneType		[ O ]�g�ѓd�b�̎��
	'				I ��i-mode
	'				J ��J-Sky
	'				E ��EzWeb
	'				P ��PC
	'				�󁨕s��

	Dim str			' ���[�U�G�[�W�F���g
	
	' ���[�U�G�[�W�F���g���擾
	str = Request.ServerVariables("HTTP_USER_AGENT")
	' ���[�U�G�[�W�F���g���ɏ���
	If Instr(1, str, "DoCoMo", 1) > 0 Then
		' ���[�U�G�[�W�F���g��i-mode�̏ꍇ
		GetPhoneType = "I"
		Exit Function
	End If

	If Instr(1, str, "J-PHONE", 1) > 0 Then
		' ���[�U�G�[�W�F���g��J-SKY�̏ꍇ
		GetPhoneType = "J"
		Exit Function
	End If

    '2010/03/18 Add-S NICS okimi	�g�т̒ǉ�
	If Instr(1, str, "Vodafone", 1) > 0 Then
		' ���[�U�G�[�W�F���g��Vodafone�̏ꍇ
		GetPhoneType = "J"
		Exit Function
	End If

	If Instr(1, str, "SoftBank", 1) > 0 Then
		' ���[�U�G�[�W�F���g���\�t�g�o���N�̏ꍇ
		GetPhoneType = "J"
		Exit Function
	End If
    '2010/03/18 Add-E NICS okimi

	If Instr(1, str, "Mozilla", 1) > 0 Then
		' ���[�U�G�[�W�F���g��PC�̏ꍇ
		GetPhoneType = "P"
		Exit Function
	End If

	If Instr(1, str, "UP.Browser", 1) > 0 Then
		' ���[�U�G�[�W�F���g��EZ-WEB�̏ꍇ
		GetPhoneType = "E"
		Exit Function
	End If

	GetPhoneType = ""
End Function

' �g�ѓd�b�̎�ʂɉ�����Title�^�O�̓��e���擾
Function GetTitleTag(strTitle)
	'stTitle		[ I ]�^�C�g��
	'GetTitleTag	[ O ]�^�C�g���^�O
	
	' J-Sky�̓^�C�g���^�O����ʂɕ\�������̂Ŏw�肵�Ȃ�
	If GetPhoneType() = "J" Then
		GetTitleTag = ""
	Else
		GetTitleTag = "<title>" & strTitle & "</title>"
	End If
End Function

' �g�ѓd�b�̎�ʂɉ������摜�̊g���q���擾
Function GetImageExt()
	'GetImageExt	[ O ]�摜�̊g���q(.�t)
	
	Select Case GetPhoneType()
		Case "I":
			GetImageExt = ".gif"
		Case "J":
			GetImageExt = ".png"
		Case "E":
			GetImageExt = ".bmp"
		Case "P":
			GetImageExt = ".gif"
		Case Else:
			GetImageExt = ".gif"
	End Select
End Function

' �g�ѓd�b�̎�ʂɉ����������{�^�������N�p�̃^�O���擾
Function GetKeyTag(strKeyName)
	'strKeyName		[ I ]�{�^����
	'GetKeyTag		[ O ]�����{�^�������N�p�̃^�O
	
	Select Case GetPhoneType()
		Case "I":
			GetKeyTag = "accesskey=""" & Right(strKeyName,1) & """"
		Case "J":
			GetKeyTag = "directkey=""" & Right(strKeyName,1) & """"
		Case "E":
			GetKeyTag = "accesskey=""" & Right(strKeyName,1) & """"
		Case Else:
			GetKeyTag = ""
	End Select
End Function

' �g�ѓd�b�̎�ʂɉ����������{�^���̃��x�����擾
Function GetKeyLabel(strKeyName)
	'strKeyName		[ I ]�{�^����
	'GetKeyLabel	[ O ]�����{�^�������N�p�̃��x��
	
	Select Case GetPhoneType()
		Case "I":
'	NEC�̋@�킪�Ή����Ă��Ȃ��̂ŊG�����̎g�p�𒆎~ 2002/02/14
'			If IsNumeric(strKeyName) Then
'				GetKeyLabel = "&#" & CStr(63878 + CInt(strKeyName))
'			Else
				GetKeyLabel = Right(strKeyName,1) & "."
'			End If
		Case "P":
			GetKeyLabel = "[" & Right(strKeyName,1) & "]"
		Case Else:
			GetKeyLabel = ""
	End Select
End Function

' �g�ѓd�b�̎�ʂɉ������e�L�X�g�{�b�N�X�̃T�C�Y�E������̃^�O���擾
Function GetTextSizeMode(nSize, strMode)
	'nSize				[ I ]�T�C�Y
	'strMode			[ I ]������(H:����, K:�J�i, A:�p��, N:����)
	'GetTextSizeMode	[ O ]�T�C�Y�E������̃^�O
	
	Dim sModeName
	Select Case GetPhoneType()
		Case "I":
			Select Case strMode
				Case "H":
					 sModeName = "1"
				Case "K":
					 sModeName = "2"
				Case "A":
					 sModeName = "3"
				Case "N":
					 sModeName = "4"
			End Select
			GetTextSizeMode = " size=""" & CStr(nSize) & """ istyle=""" & sModeName & """ "  
		Case "J":
			Select Case strMode
				Case "H":
					 sModeName = "hiragana"
				Case "K":
					 sModeName = "katakana"
				Case "A":
					 sModeName = "alphabet"
				Case "N":
					 sModeName = "numeric"
			End Select
			GetTextSizeMode = " size=""" & CStr(nSize) & """ mode=""" & sModeName & """ "
		Case Else:
			GetTextSizeMode = " size=""" & CStr(Int(nSize * 1.1 + 3)) & """ " 
	End Select
End Function

' ���o�\��(�A���R���e�i�e�[�u�����画�f)
Function CanCarryOut(rs) 
	'rs				[ I ]�A���R���e�i�e�[�u���̃��R�|�h�Z�b�g
	'CanCarryOut	[ O ]���o�\��('Y':��,'S':��,' ':�s��)

	Dim tNow
	tNow = Now

    '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r���� 20040224 S
    Dim strchkNow, strchkFreeTime, strchkDelPermitDate
    Dim strchkOLTDateFrom, strchkOLTDateTo, strchkFreeTimeExt
    strchkNow = DispDateTime( Now, 8 )
    strchkFreeTime = DispDateTime( rs("FreeTime"), 8 )
    strchkDelPermitDate = DispDateTime( rs("DelPermitDate"), 8 )
    '���o�\����NULL�̂Ƃ����o�\�ɂȂ�̂�h��(2004/09/01 Add)
    If strchkDelPermitDate = "" Then
        strchkDelPermitDate = "99999999"
    End If
    '(2004/09/01 Add End)
    strchkOLTDateFrom = DispDateTime( rs("OLTDateFrom"), 8 )
    strchkOLTDateTo = DispDateTime( rs("OLTDateTo"), 8 )
    strchkFreeTimeExt = DispDateTime( rs("FreeTimeExt"), 8 )
    '���t�܂ł��������Ă��Ȃ�DateTime�^�̔�r���� 20040224 E

    '--- add by mes(2005/3/28)
    ' �I���h�b�N�ŉ����X�g�ΏۊO�t���O���ΏۊO�łȂ��Ȃ牵���X�g�ԍ��̗L�����`�F�b�N
    If Trim(rs("OffDockFlag"))="N" And (Trim(rs("DsListFlg"))<>"1" Or IsNull(Trim(rs("DsListFlg"))) = True ) Then
    	If Trim(rs("ListNo"))="" Then
    		CanCarryOut = " "
			Exit Function
    	End If
    End If
    '--- end mes

	' �ʊց^�ېŗA���̃`�F�b�N
	If Trim(rs("OLTICFlag")) = "I" Then
		If Trim(rs("OLTICNo")) = "" Then	
'			CanCarryOut = False
			CanCarryOut = " "
			Exit Function
		End If
	'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18 Chenge)
'   Else
	ElseIf Trim(rs("OLTICFlag"))<>"" Then
	'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18)
'		If Not (tNow >= rs("OLTDateFrom") And tNow <= rs("OLTDateTo")) Then	'D20040224
		If Not (strchkNow >= strchkOLTDateFrom And strchkNow <= strchkOLTDateTo) Then 'I20040224
'			CanCarryOut = False
			CanCarryOut = " "
			Exit Function
		End If		
	'OLTICFlag���󔒂̂Ƃ��A�����Ƌ��ԍ��̃`�F�b�N�����A�ʊ�OK�Ƃ���(2005/02/03 mod)
    Else
		If DispDateTime( rs("OLTICDate"), 8 ) <> "" And Trim(rs("OLTICNo")) = "" Then
			CanCarryOut = "Y"
		End If
	'OLTICFlag���󔒂̂Ƃ��A���t�`�F�b�N���s�킸�A�ʊ�OK�Ƃ���(2005/01/18)
	End If

'	If Not IsNull(rs("CYDelTime")) Then
	If DispDateTime(rs("CYDelTime"),0)<>"" Then
		CanCarryOut = "S"
'	ElseIf Trim(rs("DOStatus")) = "Y" And _
'		(rs("FreeTime") >= tNow Or rs("FreeTimeExt") >= tNow) And _
'		rs("DelPermitDate") <= tNow Then	'D20040224
'--- mod by MES(2004/9/10)
'	ElseIf Trim(rs("DOStatus")) = "Y" And _
'		(strchkFreeTime >= strchkNow Or strchkFreeTimeExt >= strchkNow) And _
'		strchkDelPermitDate <= strchkNow Then	'I20040224
	ElseIf Trim(rs("DOStatus")) = "Y" And _
		(strchkFreeTime >= strchkNow Or strchkFreeTimeExt >= strchkNow) And _
		strchkDelPermitDate <= strchkNow AND _
		Trim(rs("OLTICNo")) <> "" AND DispDateTime( rs("OLTICDate"), 8 ) <> "" Then
'--- end MES
		
		CanCarryOut = "Y"
	Else

		CanCarryOut = " "
	End If
End Function

' ���l���w�肵�������̕�����ɕϊ�(�E�l�E�]���ɂ͂O)
Function ArrangeNum(nNumber, nFigure)
	'nNumber	[ I ]�ϊ����鐔�l
	'nFigure	[ I ]�ϊ���̌���
	'ArrangeNum	[ O ]�ϊ���̕�����

	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNum = sNum
End Function

'���ʊm�F�p�f�[�^�̏o��(�g�ѓd�b�p)
Function WriteLogM(fs, strUserID, iNum, strScrName, iWrkNum, strPhoneKind,strOption)
	'fs			[ I ]�t�@�C���V�X�e��Object
	'strUserID	[ I ]���[�UID
	'iNum		[ I ]���ID
	'strScrName	[ I ]��ʖ�
	'iWrkNum	[ I ]���ID
	'strPhoneKind [ I ]�g�ы@��
	'strOption	[ I ]�m�F�p�I�v�V�����f�[�^


	' Log���o�͂���t�H���_
	Const LOG_FOLDER = "log"

	' Log�t�@�C���̊g���q 
	Const LOG_EXTENT = "log"

	Dim strDate,strNowTime
	Dim logPath
	Dim logs

	strNowTime = Year(Now) & "/" & _
		Right("0" & Month(Now), 2) & "/" & _
		Right("0" & Day(Now), 2) & " " & _
		Right("0" & Hour(Now), 2) & ":" & _
		Right("0" & Minute(Now), 2) & ":" & _
		Right("0" & Second(Now), 2)
	strDate = GetYMDStr(Now)
	logPath = Server.MapPath(LOG_FOLDER & "/" & strDate & "." & LOG_EXTENT)
	Set logs=fs.OpenTextFile(logPath, 8, True)
	logs.Write strNowTime
    logs.Write "," & iNum
	logs.Write "," & strScrName
    logs.Write "," & iWrkNum
	If strPhoneKind="E" Then
		logs.Write ",A"
	Else
		logs.Write "," & strPhoneKind
	End If
	logs.Write "," & strUserID
	logs.Write "," & Request.ServerVariables("SCRIPT_NAME")
'	logs.Write "," & Request.ServerVariables("CONTENT_LENGTH")
	logs.Write "," & Request.ServerVariables("REMOTE_ADDR")
	If strOption<>"" Then
		logs.WriteLine "," & strOption
	Else
		logs.WriteLine 
	End If

	logs.Close

    '2009/03/05 Add-S Tanaka
     xLogReg iNum,iWrkNum,strUserID
    '2009/03/05 Add-E Tanaka

End Function

'2009/03/05 Add-S Tanaka
'���O��DB�ɓo�^
Function xLogReg(iNum,iWrkNum,strUserId)
	'iNum		[ I ]���ID
	'iWrkNum	[ I ]���ID

	Dim ObjConn, ObjRS, StrSQL,fs,ti,strTemp,anyTmpTitle
	Dim iAddFlag
	Dim PageNum(),WkNum(),PageTitle(),SubTitle()
	Dim strTitleFileName

	'DB�ڑ�
	ConnectSvr ObjConn,ObjRS


	Set fs=Server.CreateObject("Scripting.FileSystemObject")

	' ���O�^�C�g���擾
	strTitleFileName="../logija.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iAddFlag=0
	
	'���O�^�C�g�������[�v
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")

		'���O�^�C�g���ɊY��������̏ꍇ
		IF TRIM(anyTmpTitle(0))=TRIM(CSTR(iNum)) and TRIM(anyTmpTitle(1))=TRIM(CSTR(iWrkNum)) then
			'���O�C���R�[�h���Y���̂���
			IF strUserId<>"90000" and strUserId<>"10000" and strUserId<>"MES" and strUserId<>"MEO" Then
				iAddFlag=1
				Exit Do
			ELSE
				Exit Function
			End If
		End If
	Loop
	ti.Close

	'�o�^�\�ȃf�[�^�����݂���ꍇ
	If iAddFlag=1 Then
		StrSQL = ""
		StrSQL = StrSQL & " SELECT * FROM xLog "
		StrSQL = StrSQL & " WHERE LogDate='" & GetYMDStr(Now) &"'"
		StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
		StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		
		ObjRS.Open StrSQL, ObjConn

		'���R�[�h�����݂���ꍇ
		If ObjRS.EOF=False then
			StrSQL = ""
			StrSQL = StrSQL & " UPDATE xLog SET "
			StrSQL = StrSQL & " DataCount=DataCount+1 "
			StrSQL = StrSQL & " ,UpdtTime=GetDate() "
			StrSQL = StrSQL & " ,UpdtPgCd='MCOMNINC' "
			StrSQL = StrSQL & " ,UpdtTmnl='" & Session.Contents("userid") & "'"
			StrSQL = StrSQL & " WHERE LogDate='" & GetYMDStr(Now) &"'"
			StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
			StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		ELSE
			StrSQL = ""
			StrSQL = StrSQL & " INSERT INTO xLog( "
			StrSQL = StrSQL & " LogDate, "
			StrSQL = StrSQL & " PageNum, "
			StrSQL = StrSQL & " WkNum, "
			StrSQL = StrSQL & " UpdtPgCd, "
			StrSQL = StrSQL & " UpdtTime, "
			StrSQL = StrSQL & " UpdtTmnl, "
			StrSQL = StrSQL & " DataCount "
			StrSQL = StrSQL & " )VALUES( "
			StrSQL = StrSQL & "'" & GetYMDStr(Now) & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iNum)) & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iWrkNum)) & "',"
			StrSQL = StrSQL & "'MCOMNINC',"
			StrSQL = StrSQL & "GetDate(),"
			StrSQL = StrSQL & "'" & Session.Contents("userid") & "',"
			StrSQL = StrSQL & "1)"
		End If
		ObjConn.Execute(StrSQL)
	End If

End Function
'2009/03/05 Add-E Tanaka


'���t�̕\���P
Function DispDateTime(vDate,length)
'	vDate			[ I ]���t
	Dim strTmp
    If VarType(vDate)<>vbDate Then
        strTmp = ""
    Else
        strTmp = _
            Year(vDate) & "/" & _
            Right("0" & Month(vDate), 2) & "/" & _
            Right("0" & Day(vDate), 2) & " " & _
            Right("0" & Hour(vDate), 2) & ":" & _
            Right("0" & Minute(vDate), 2)
        If length=11 Then
            strTmp = Right(strTmp, 11)
        ElseIf length=10 Then
            strTmp = Left(strTmp, 10)
        ElseIf length=8 Then
            strTmp = Mid(strTmp, 3, 8)
        ElseIf length=5 Then
            strTmp = Mid(strTmp, 6, 5)
        End If
        DispDateTime = strTmp
    End If
End Function

%>