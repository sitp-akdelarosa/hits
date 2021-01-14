<% 
'�c�a�ڑ�
Function ConnectSvr(conn, rsd) 
	'conn		[ O ]DB�ڑ�Object
	'rsd		[ O ]���R�[�h�Z�b�gObject

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "Provider=SQLOLEDB; Data Source = SI-KURT-LAPTOP\SQLEXPRESS; Initial Catalog = CIDB; User Id = sa; Password=P@55w0rd"
'	conn.Open "KAIRIKU_TEST", "sa", "icct"

	Set rsd = Server.CreateObject("ADODB.Recordset")
End Function

'���ʊm�F�p�f�[�^�̏o��
Function WriteLog(fs, strScrName, strOption)
	'fs			[ I ]�t�@�C���V�X�e��Object
	'strScrName	[ I ]��ʖ�
	'strOption	[ I ]�m�F�p�I�v�V�����f�[�^

    strDate = GetYMDStr(Now)
    logPath = Server.MapPath("log/" & strDate & ".log")
    Set logs=fs.OpenTextFile(logPath, 8, True)
    logs.Write Now
    logs.Write Chr(9) & strScrName
    logs.Write Chr(9) & Session.Contents("userkind") & ":" & Session.Contents("userid")
    logs.Write Chr(9) & Request.ServerVariables("SCRIPT_NAME")
'   logs.Write Chr(9) & Request.ServerVariables("CONTENT_LENGTH")
    logs.Write Chr(9) & Request.ServerVariables("REMOTE_ADDR")
    If strOption<>"" Then
        logs.WriteLine Chr(9) & strOption
    Else
        logs.WriteLine 
    End If

    logs.Close
End Function

'���[�U�[�h�c���L�����ǂ����̃`�F�b�N
Function CheckLogin(strReturn)
	'strReturn	[ I ]���[�U�[�h�c�����̂Ƃ��̖߂����

    ' �Z�b�V�����̗L�������`�F�b�N
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))

    ' �Z�b�V�����������ȂƂ�
    If strUserID="" Then
        ' �߂��ʂփ��_�C���N�g
        Response.Redirect "userchk.asp?link=" & strReturn
    End If
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

'�Œ茅���ł̕\��
Function GetNumStr(vNum, iK)
'	vNum			[ I ]���l�i���lor������j
'	iK				[ I ]����
	GetNumStr = Right(String(iK, "0") & CStr(vNum), iK)
End Function

'���t�̕\���P
Function DispDateTime(vDate,length)
'	vDate			[ I ]���t
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

'���t�̕\���Q
Function DispDateTimeCell(sCell,length)
'	sCell			[ I ]���t
    If VarType(sCell)<>vbString Or sCell="" Then
        DispDateTimeCell = "<hr width=80%" & ">"
    Else
        If length=5 Then
            strTmp = Mid(sCell, 6, 5)
        ElseIf length=10 Then
            strTmp = Mid(sCell, 6, 5) & "<br>" & Right(sCell, 5)
        ElseIf length=11 Then
            strTmp = Right(sCell, 11)
        Else
            strTmp = sCell
        End If
        DispDateTimeCell = strTmp
    End If
End Function

'�\����̕\��
Function DispReserveCell(sFlag, sCell, sColor)
'	sFlag			[ I ]�t���O
'	sCell			[ I ]���t
'	sColor			[ O ]�F
    If sFlag="" Or sFlag="0" Then
        DispReserveCell = "<hr width=80%" & ">"
        sColor=""
    Else
        If sFlag="1" Then
            strTmp = Mid(sCell, 6, 11)
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="2" Then
            strTmp = Mid(sCell, 6, 5) & " �ߑO"
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="3" Then
            strTmp = Mid(sCell, 6, 5) & " �ߌ�"
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="4" Then
            strTmp = Mid(sCell, 6, 5) & " �[��"
            sColor="<font color='#0000FF'>"
        ElseIf sFlag="5" Then
            strTmp = "�\��폜"
            sColor="<font color='#FF0000'>"
        End If
        DispReserveCell = strTmp
    End If
End Function

'�t���O�̕\��
Function DispFlag(vFlg)
'	vFlg			[ I ]�t���O������
    If vFlg=vbNull Or Trim(vFlg)="" Then
        DispFlag = ""
    Else
        If vFlg="1" Then
            DispFlag = "�~"
        ElseIf vFlg="2" Then
            DispFlag = "��"
        End If
    End If
End Function

'�\�̕\��
Function DispTableCell(sCell,i)
'	sCell			[ I ]������
    If VarType(sCell)<>vbString Or sCell="" Then
        If i=1 Then
            DispTableCell = "<hr width=80%" & ">"
        Else
            DispTableCell = "<br>"
        End If
    Else
        DispTableCell = sCell
    End If
End Function

'Menu�I���R���{�{�b�N�X��Java�֐�
Sub DispMenuJava
    Response.Write "<!--" & Chr(13) & Chr(10)
    Response.Write "function LinkSelect(form, sel)" & Chr(13) & Chr(10)
    Response.Write "{" & Chr(13) & Chr(10)
    Response.Write "    adrs = sel.options[sel.selectedIndex].value;" & Chr(13) & Chr(10)
    Response.Write "    if (adrs != '-' ) parent.location.href = adrs;" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)

	'�R�[�h�ꗗ�pfunction�ǉ�	2002/02/06
    Response.Write "function OpenCodeWin()" & Chr(13) & Chr(10)
    Response.Write "{" & Chr(13) & Chr(10)
    Response.Write "	window.open(""codelist.asp"",""codelist"",""scrollbars=yes,resizable=yes,width=250,height=270"");" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)

    Response.Write "// -->" & Chr(13) & Chr(10)
End Sub

'Menu�I���R���{�{�b�N�X�̕`��
Sub DispMenu
    Response.Write "<FORM action=''>"

	'�R�[�h�ꗗ�p�{�^���ǉ�	2002/02/06
    Response.Write "<input type=button value="" �R�[�h�ꗗ "" onclick=""JavaScript:OpenCodeWin()"">"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

    Response.Write "<SELECT NAME='link' onchange='LinkSelect(this.form, this)'>"
    Response.Write "<OPTION VALUE='#'>Contents"
    Response.Write "<option value='index.asp'>TOP</option>"
    Response.Write "<option value='#'>�R���e�i���Ɖ� </option>"
    Response.Write "<option value='userchk.asp?link=expentry.asp'>�� �A�o�R���e�i���Ɖ� </option>"
    Response.Write "<option value='userchk.asp?link=impentry.asp'>�� �A���R���e�i���Ɖ� </option>"
    Response.Write "<option value='#'>�e�Г��͉��</option>"
    Response.Write "<option value='userchk.asp?link=nyuryoku-in1.asp'>�� �D��/�^�[�~�i������ "
    Response.Write "</option>"
    Response.Write "<option value='userchk.asp?link=nyuryoku-kaika.asp'>�� �C�ݓ��� </option>"
    Response.Write "<option value='userchk.asp?link=nyuryoku-te.asp'>�� �^�[�~�i������ </option>"
    Response.Write "<option value='userchk.asp?link=rikuun1.asp'>�� ���^����</option>"
    Response.Write "<option value='userchk.asp?link=sokuji.asp'> �������o�V�X�e�� </option>"
    Response.Write "<option value='userchk.asp?link=hits.asp'>�X�g�b�N���[�h���p�V�X�e��</option>"
    Response.Write "<option value='userchk.asp?link=terminal.asp'>�Q�[�g�O�f���E���G�󋵏Ɖ� </option>"
    Response.Write "<option value='userchk.asp?link=request.asp'>���p�҃A���P�[�g�E�p���`</option>"
    Response.Write "</SELECT>"
    Response.Write "</FORM>"
End Sub

'Menu�o�[�̕`��
Sub DispMenuBar
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td valign='bottom' align='right'><a href='index.asp'><img src='gif/b-home.gif' border='0' width='270' height='23' usemap='#map'></a></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td bgcolor='000099' height='10'><img src='gif/1.gif'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

'�߂��̓o�^
Sub DispMenuBarBack(strBackURL)
'    ' �߂���o�^
'    Dim strBackURL
'    strBackURL = Trim(Session.Contents("returnurl"))
'    If strBackURL="" Then
'        strBackURL = "index.asp"
'    End If

    Response.Write "<map name='map'>"
    Response.Write "<area shape='poly' coords='20,0,152,0,134,22,0,22' href='" & strBackURL & "'>"
    Response.Write "<area shape='poly' coords='154,0,136,22,284,22,284,0' href='index.asp'>"
    Response.Write "</map>"

'    Session.Contents("returnurl") = Request.ServerVariables("SCRIPT_NAME")
End Sub

'�G���[���b�Z�[�W�̕\��
Sub DispErrorMessage(strMessage)
    Response.Write "<center>"
    Response.Write "<table border=1 cellpadding='2' cellspacing='1'>"
    Response.Write "<tr>"
    Response.Write "<td bgcolor='#FFFFFF'>"
    Response.Write "<table border='0'>"
    Response.Write "<tr>"
    Response.Write "<td valign='middle'><img src='gif/error.gif'></td>"
    Response.Write "<td><b><font color='#FF4400'>"
    Response.Write strMessage
    Response.Write "</font></b></td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</center>"
End Sub

'���ڃ��b�Z�[�W�̕\��
Sub DispInformationMessage(strMessage)
    Response.Write "<center>"
    Response.Write "<table border=1 cellpadding='2' cellspacing='1'>"
    Response.Write "<tr>"
    Response.Write "<td bgcolor='#FFFFFF'>"
    Response.Write "<table border='0'>"
    Response.Write "<tr>"
    Response.Write "<td valign='middle'><img src='gif/info.gif'></td>"
    Response.Write "<td><b><font color='#0000FF'>"
    Response.Write strMessage
    Response.Write "</font></b></td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</center>"
End Sub

' �g�����U�N�V�����@�����V�[�P���X�ԍ����s
Function GetDailyTransNo()
	'�߂�l		[ O ]�����V�[�P���X�ԍ�(������)

	' ����
	Const FIGURE_NUM = 6

	If Application("LastTransDate") <> Date Or Application("DailyTransNo") = "" Then
		' �V�X�e���N����܂��͓��t���ς�����ꍇ�͂P��ݒ�
		Application("LastTransDate") = Date
		Application("DailyTransNo") = 1
	Else
		' �ʏ�͂P�����Z
		Application("DailyTransNo") = Application("DailyTransNo") + 1
	End If 

	' ���߂�ꂽ�����̕������
	Dim sNo

	sNo = CStr(Application("DailyTransNo"))
	If FIGURE_NUM < Len(sNo) Then
		sNo = Right(sNo, FIGURE_NUM)
	Else
		Do While Len(sNo) < FIGURE_NUM 
			sNo = "0" & sNo
		Loop				
	End If

	GetDailyTransNo = sNo
End Function

'Temp�t�@�C���̎�ނ�ݒ�
Sub SetTempFile(strKind)
    Session.Contents("tempfilekind")=strKind
End Sub

'Temp�t�@�C�����L�����ǂ����̃`�F�b�N
Function CheckTempFile(strKind,strReturn)
	'strKind	[ I ]Temp�t�@�C���̎��
	'strReturn	[ I ]Temp�t�@�C�������̂Ƃ��̖߂����

    ' Temp�t�@�C���̗L�������`�F�b�N
    Dim strTempKind
    strTempKind = Trim(Session.Contents("tempfilekind"))

    ' Temp�t�@�C���������ȂƂ�
    If strTempKind<>strKind Then
        ' �߂��ʂփ��_�C���N�g
        Response.Redirect strReturn
    End If
End Function

'�����Q���ϊ�
Function DateFormat(strNum)
	If strNum="1" Then
		strNum = "01"
	End If
	If strNum="2" Then
		strNum = "02"
	End If
	If strNum="3" Then
		strNum = "03"
	End If
	If strNum="4" Then
		strNum = "04"
	End If
	If strNum="5" Then
		strNum = "05"
	End If
	If strNum="6" Then
		strNum = "06"
	End If
	If strNum="7" Then
		strNum = "07"
	End If
	If strNum="8" Then
		strNum = "08"
	End If
	If strNum="9" Then
		strNum = "09"
	End If
    DateFormat = strNum
End Function

'/*
' *	�ꏊ�\���p�̕������Ԃ�
' *
' *	@param	String	strScriptName	ASP�t�@�C����
' *	@return	String	strRoute		Top���猻�݂̃y�[�W�܂ł̓���
' */
Function SetRoute(strScriptName)
	Dim bRouteFlag,iIndex,strParent,strSubName,strRoute,strExpCont,strImpCont,strTerminal

	iIndex = InStrRev(strScriptName,"/")
	strScriptName = Right(strScriptName,Len(strScriptName)-iIndex)

	strExpCont 	= "�A�o�R���e�i���Ɖ�i��ƑI���j "
	strImpCont 	= "�A���R���e�i���Ɖ�i��ƑI���j "
	strTerminal = "�^�[�~�i�����p���ԏƉ� "
	strKaika 	= "�C�ݓ��́i��ƑI���j "
	strVslTerm 	= "�D�Ё^�^�[�~�i�����́i��ƑI���j "
	strTerm 	= "�Q�[�g�O�f���E���G�󋵏Ɖ� "
	strTrack 	= "���^���́i�R���e�iNo.���́j "
	strGate 	= "�Q�[�g�ʍs���ԗ\��V�X�e�� "
	strReq	 	= "���p�҃A���P�[�g "

	bRouteFlag = true
	strParent  = ""

    Select Case strScriptName

	'�A�o�R���e�i���Ɖ�
        Case "expentry.asp","expcntnr.asp","expreload.asp"
								strSubName = strExpCont
								bRouteFlag = false
        Case "expcsv.asp","expcsvin.asp"
								strSubName = "�A�o�R���e�i���Ɖ�p�t�@�C���]����� "
								strParent  = strExpCont
		Case "expdetail.asp"	strSubName = "�A�o�R���e�i�Ɖ�i�P�ƃR���e�i�j "
        Case "explist.asp"		strSubName = "�A�o�R���e�i�Ɖ�i�����R���e�i�j "
								strParent  = strExpCont
        Case "explist1.asp"		strSubName = "�ʒu��� "
								strParent  = "�A�o�R���e�i�Ɖ�i�����R���e�i�j "
        Case "explist2.asp"		strSubName = "�R���e�i�쐬�ɌW���� "
								strParent  = "�A�o�R���e�i�Ɖ�i�����R���e�i�j "
        Case "explist3.asp"		strSubName = "�^�[�~�i���A�{�D�ɌW���� "
								strParent  = "�A�o�R���e�i�Ɖ�i�����R���e�i�j "

		'�A�o�R���e�i���Ɖ�i�����d�l�j
        Case "ms-expentry.asp"
								strKind = Request.QueryString("kind")
								Select Case strKind
									Case 1	strSubName = "�C�ݏƉ� "
									Case 2	strSubName = "���^�Ɖ� "
									Case 3	strSubName = "�׎�Ɖ� "
								End Select
								strParent  = strExpCont
        Case "ms-explist1.asp","ms-explist2.asp","ms-explist3.asp"
								strSubName = "�A�o�R���e�i���ꗗ "
        Case "ms-expdetail.asp"	strSubName = "�R���e�i���ڍ� "
								strParent  = "�A�o�R���e�i���ꗗ "
        Case "ms-expinput.asp","ms-expinput-syori.asp"
								strSubName = "�������� "

	'�A���R���e�i���Ɖ�
        Case "impentry.asp","impcntnr.asp","impreload.asp"
								strSubName = strImpCont
								bRouteFlag = false
        Case "impcsv.asp","impcsvin.asp"
								strSubName = "�A���R���e�i���Ɖ�p�t�@�C���]����� "
								strParent  = strImpCont
		Case "impdetail.asp"	strSubName = "�A���R���e�i�Ɖ�i�P�ƃR���e�i�j "
        Case "implist.asp"		strSubName = "�A���R���e�i�Ɖ�i�����R���e�i�j "
								strParent  = strImpCont
        Case "implist1.asp"		strSubName = "�^�[�~�i�������܂ł̈ʒu��� "
								strParent  = "�A���R���e�i�Ɖ�i�����R���e�i�j "
        Case "implist2.asp"		strSubName = "�^�[�~�i�����o��̊�{��񂨂�шʒu��� "
								strParent  = "�A���R���e�i�Ɖ�i�����R���e�i�j "

		'�A���R���e�i���Ɖ�i�����d�l�j
        Case "ms-impentry.asp"
								strKind = Request.QueryString("kind")
								Select Case strKind
									Case 1	strSubName = "�C�ݏƉ� "
									Case 2	strSubName = "���^�Ɖ� "
									Case 3	strSubName = "�׎�Ɖ� "
								End Select
								strParent  = strImpCont
        Case "ms-implist1.asp","ms-implist2.asp"
								strSubName = "�A���R���e�i���ꗗ "
        Case "ms-impdetail.asp"	strSubName = "�R���e�i���ڍ� "
								strParent  = "�A���R���e�i���ꗗ "
        Case "ms-impinput.asp","ms-impinput-syori.asp"
								strSubName = "�������� "

	'�^�[�~�i�����p���ԏƉ�
	'�^�[�~�i�����p���ԏƉ�
        Case "terminal.asp"		strSubName = strTerm
								bRouteFlag = false
        Case "camera.asp"		strSubName = "�Q�[�g���Ӊf�� "
								strParent  = strTerm
        Case "photogate.asp"	strSubName = "�Q�[�g�O�f�� "
								strParent  = "�Q�[�g���Ӊf�� "
        Case "phototaikijyo.asp"	strSubName = "�ҋ@��f�� "
									strParent  = "�Q�[�g���Ӊf�� "
        Case "photokamome.asp"	strSubName = "�����ߑ勴�f�� "
								strParent  = "�Q�[�g���Ӊf�� "

	'���ʃ��O�C�����
        Case "userchk.asp"
							If InStr(Request.QueryString("link"),"-impentry.asp")<>0 Then
								strSubName = strImpCont & "&gt; ���O�C��"
								Session.Contents("route") = "Top &gt; " & strImpCont
							ElseIf InStr(Request.QueryString("link"),"-expentry.asp")<>0 Then
								strSubName = strExpCont & "&gt; ���O�C��"
								Session.Contents("route") = "Top &gt; " & strExpCont
							Else
								strSubName = "���O�C��"
								Session.Contents("route") = ""
							End If
							bRouteFlag = false

	'�D�Ё^�^�[�~�i������
        Case "nyuryoku-in1.asp"				strSubName = strVslTerm
											bRouteFlag = false
        Case "nyuryoku-csv.asp","nyuryoku-csvin.asp"
											strSubName = "�D�Ё^�^�[�~�i���p�t�@�C���]����� "
        Case "nyuryoku-port.asp"			strSubName = "�{�D���Èꗗ "
        Case "nyuryoku-vsl.asp"				strSubName = ""
        Case "nyuryoku-sch.asp"				strSubName = "�{�D���Ó��́i�X�V�j"
        Case "nyuryoku-new.asp"				strSubName = "�{�D���Ó��́i�V�K�j"
        Case "nyuryoku-new-ist.asp"			strSubName = ""

	'�C�ݓ���
        Case "nyuryoku-kaika.asp"			strSubName = strKaika
											bRouteFlag = false
		'�b�x�������w��
        Case "nyuryoku-ex.asp"				strSubName = "�A�o�b�x�������w�� "
        Case "nyuryoku-ex-csv.asp","nyuryoku-ex-csvin.asp"
											strSubName = "�A�o�b�x�������w���p�t�@�C���]����� "
        Case "nyuryoku-ex-syori.asp"		strSubName = ""

		'�V�[���ԍ��E�d��
        Case "nyuryoku-ki.asp"				strSubName = "�A�o�V�[���ԍ��E�d�� "
        Case "nyuryoku-kcsv.asp","nyuryoku-kcsvin.asp"
											strSubName = "�A�o�V�[���ԍ��E�d�ʗp�t�@�C���]����� "
        Case "nyuryoku-ki-syori.asp"		strSubName = ""

		'������q�ɓ͎����w��
        Case "nyuryoku-im.asp"				strSubName = "�A��������q�ɓ͂������w�� "
        Case "nyuryoku-im-csv.asp","nyuryoku-im-csvin.asp"
											strSubName = "�A��������q�ɓ͂������w���p�t�@�C���]����� "
        Case "nyuryoku-im-syori.asp"		strSubName = ""

		'�A�o�ݕ����
        Case "ms-kaika-expinfo.asp"			strSubName = "�A�o�ݕ������� "
        Case "ms-kaika-expinfo-new.asp"
											If Request.QueryString("kind")<>0 Then
												strSubName = "�V�K���� "
											Else
												strSubName = "�X�V���� "
											End If
        Case "ms-kaika-expinfo-list.asp"	strSubName = "�X�V�Ώۈꗗ "
        Case "ms-kaika-expinfo-csv.asp","ms-kaika-expinfo-csvin.asp"
											strSubName = "�ݕ������͗p�t�@�C���]����� "
        Case "ms-kaika-expinfo-exec.asp","ms-kaika-expinfo-updatecheck.asp"
											strSubName = ""

		'�A�o�R���e�i���
        Case "ms-kaika-expcontinfo.asp"		strSubName = "�A�o�R���e�i������ "
        Case "ms-kaika-expcontinfo-new.asp"
											If Request.QueryString("kind")<>0 Then
												strSubName = "�V�K���� "
											Else
												strSubName = "�X�V���� "
											End If
        Case "ms-kaika-expcontinfo-list.asp"	strSubName = "�X�V�Ώۈꗗ "
        Case "ms-kaika-expcontinfo-csv.asp","ms-kaika-expcontinfo-csvin.asp"
												strSubName = "�R���e�i�����͗p�t�@�C���]����� "
        Case "ms-kaika-expcontinfo-exec.asp","ms-kaika-expcontinfo-updatecheck.asp"
												strSubName = ""

		'�A���R���e�i���
        Case "ms-kaika-impcontinfo.asp"		strSubName = "�A���R���e�i������ "
        Case "ms-kaika-impcontinfo-new.asp"
											If Request.QueryString("kind")<>0 Then
												strSubName = "�V�K���� "
											Else
												strSubName = "�X�V���� "
											End If
        Case "ms-kaika-impcontinfo-list.asp"	strSubName = "�X�V�Ώۈꗗ "
        Case "ms-kaika-impcontinfo-csv.asp","ms-kaika-impcontinfo-csvin.asp"
												strSubName = "�R���e�i�����͗p�t�@�C���]����� "
        Case "ms-kaika-impcontinfo-exec.asp","ms-kaika-impcontinfo-updatecheck.asp"
												strSubName = ""

	'�^�[�~�i������
		Case "nyuryoku-te.asp"				strSubName = strTerm
											bRouteFlag = false
		Case "nyuryoku-te1.asp"
											If strChoice="bl" Then
												strSubName = "�����m�F�\�莞������( BL�P�� ) "
											Else
												strSubName = "�����m�F�\�莞������( �{�D�P�� ) "
											End If
		Case "nyuryoku-tmnl-csv.asp","nyuryoku-tmnl-csvin.asp"
											strSubName = "�����m�F�\�莞�����͗p�t�@�C���]����� "

	'���^����
		Case "rikuun1.asp"					strSubName = strTrack
											bRouteFlag = false
		Case "rikuun2.asp","rikuun3.asp"	strSubName = "������Ƒ��M "
											strParent  = strTrack

	'�Q�[�g�ʍs���ԗ\��V�X�e��
		Case "gate.asp"						strSubName = strGate
											bRouteFlag = false

	'���p�҃A���P�[�g
		Case "request.asp"					strSubName = strReq
											bRouteFlag = false

		Case Else
	End Select


	If bRouteFlag Then

		If strSubName="" Then
			strRoute = Session.Contents("route")
		ElseIf strParent<>"" Then
			strSubName = " &gt; " & strSubName
			strRoute = Session.Contents("route")
			iIndex = InStr(strRoute,strParent)
			strRoute = Left(strRoute,iIndex-1) & strParent & strSubName
		Else
			strSubName = " &gt; " & strSubName
			strRoute = Session.Contents("route")
			iIndex = InStr(strRoute,strSubName)
			If iIndex<>0 Then
				strRoute = Left(strRoute,iIndex) & strSubName
			Else
				strRoute = strRoute &strSubName
			End If
		End If

	Else
		strRoute = "Top &gt; " & strSubName
	End If

	SetRoute = strRoute

End Function

Function GetUpdateTime(fs)

	Dim strPath,f
    strFileName = Session.Contents("tempfile")

	If strFileName<>"" Then
	    strFileName="./temp/" & strFileName
		strPath = Server.MapPath(strFileName)
	Else
		GetUpdateTime = ""
		Exit Function
	End If

	Set f = fs.GetFile(strPath)
	dateTimeTmp = f.DateLastModified
	GetUpdateTime = Year(dateTimeTmp) & "�N" & _
		Right("0" & Month(dateTimeTmp), 2) & "��" & _
		Right("0" & Day(dateTimeTmp), 2) & "��" & _
		Right("0" & Hour(dateTimeTmp), 2) & "��" & _
		Right("0" & Minute(dateTimeTmp), 2) & "�����݂̏��"

End Function

%>