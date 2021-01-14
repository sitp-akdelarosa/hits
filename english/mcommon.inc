<%
' ���[�UID�̌`���`�F�b�N
Function CheckUserID(sUserID)
	'sUserID		[ I ]���[�UID
	'CheckUserID	[ O ]�G���[���b�Z�[�W

	' �^����Љ�ЃR�[�h(���[�UID�̐擪)�̍ő�l
	Const COMPANY_MAX = "200"

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
			If IsNumeric(strKeyName) Then
				GetKeyLabel = "&#" & CStr(63878 + CInt(strKeyName))
			Else
				GetKeyLabel = "[" & Right(strKeyName,1) & "]"
			End If
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
	'CanCarryOut	[ O ]���o�\��

	Dim tNow
	tNow = Now
	
	' �ʊց^�ېŗA���̃`�F�b�N
	If Trim(rs("OLTICFlag")) = "I" Then
		If Trim(rs("OLTICNo")) = "" Then	
			CanCarryOut = False
			Exit Function
		End If
	Else
		If Not (tNow >= rs("OLTDateFrom") And tNow <= rs("OLTDateTo")) Then
			CanCarryOut = False
			Exit Function
		End If		
	End If

	If Trim(rs("DOStatus")) = "Y" And _
		(rs("FreeTime") >= tNow Or rs("FreeTimeExt") >= tNow) And _
		rs("DelPermitDate") <= tNow Then
		
		CanCarryOut = True
	Else
		CanCarryOut = False
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
Function WriteLogM(fs, strUserID, strScrName, strOption)
	'fs			[ I ]�t�@�C���V�X�e��Object
	'strUserID	[ I ]���[�UID
	'strScrName	[ I ]��ʖ�
	'strOption	[ I ]�m�F�p�I�v�V�����f�[�^

	' Log���o�͂���t�H���_
	Const LOG_FOLDER = "../log"

	' Log�t�@�C���̊g���q 
	Const LOG_EXTENT = "log"

	Dim strDate
	Dim logPath
	Dim logs

	strDate = GetYMDStr(Now)
	logPath = Server.MapPath(LOG_FOLDER & "/" & strDate & "." & LOG_EXTENT)
	Set logs=fs.OpenTextFile(logPath, 8, True)
	logs.Write Now
	logs.Write Chr(9) & strScrName
	logs.Write Chr(9) & strUserID
	logs.Write Chr(9) & Request.ServerVariables("SCRIPT_NAME")
'	logs.Write Chr(9) & Request.ServerVariables("CONTENT_LENGTH")
	logs.Write Chr(9) & Request.ServerVariables("REMOTE_ADDR")
	If strOption<>"" Then
		logs.WriteLine Chr(9) & strOption
	Else
		logs.WriteLine 
	End If

	logs.Close
End Function
%>