<% 
'�{�D���ÁE�{�D��`�n�Ɖ�
Function SerchKVsl(conn, rsd, pfp, sSeq, sDate, sName, sPlace, sVsl, sVoy, sBL, sEDate)
    'conn       [ I ]DB�ڑ�Object
    'rsd        [ I ]���R�[�h�Z�b�gObject
    'pfp        [ I ]�e���|�����t�@�C��
    'sVsl     	[ I ]�D���i�R�[���T�C���j
    'sVoy     	[ I ]Voyage�ԍ�
    'sBL    	[ I ]BL�ԍ�

    bWrite = 0        '�o�̓��R�[�h����
	bErrFlag = 0

    ' ���̓R�[���T�C���̃`�F�b�N
	ConnectSvr conn, rsd
	sql = "SELECT FullName FROM mVessel WHERE VslCode='" & sVsl & "'"

	'SQL�𔭍s����mVessel������
	rsd.Open sql, conn, 0, 1, 1
	If rsd.EOF Then
		bErrFlag = 1
	End If
	rsd.Close

	sql = "SELECT VoyCtrl FROM VslSchedule WHERE VslCode='" & sVsl & "' AND DsVoyage='" & sVoy & "'" 

	'SQL�𔭍s����VslSchedule������
	rsd.Open sql, conn, 0, 1, 1
	If rsd.EOF Then
		bErrFlag = 1
	Else
		iVoyCtrl = rsd("VoyCtrl")
	End If
	rsd.Close

	If sBL<>"" Then
		sql = "SELECT ShipLine FROM BL WHERE VslCode='" & sVsl & "' And VoyCtrl=" & iVoyCtrl & " And BLNo='" & sBL & "'"

		'SQL�𔭍s����VslSchedule������
		rsd.Open sql, conn, 0, 1, 1
		If rsd.EOF Then
			bErrFlag = 1
		End If
		rsd.Close
	End If

	If Not bErrFlag Then
		strOut = sSeq & ","
        strOut = strOut & "IM15,R,"
        strOut = strOut & sDate & ","
        strOut = strOut & sName & ","
        strOut = strOut & sPlace & ","
        strOut = strOut & sVsl & ","
        strOut = strOut & sVoy & ","
        strOut = strOut & sBL & ","
		strOut = strOut & sEDate

        pfp.WriteLine strOut
        bWrite = bWrite + 1

	End If

    SerchKVsl = bWrite    '�Y�����R�[�h���������ǂ���
End Function

'���t�̃Z�b�g
Function SetDateTime(vYear, vMonth, vDay, vHour, vMin)
'	vYear			[ I ]���t(�N)
'	vMonth			[ I ]���t(��)
'	vDay			[ I ]���t(��)
'	vHour			[ I ]���t(��)
'	vMin			[ I ]���t(��)
        SetDateTime = trim(vYear) & "/" & _
            Right("0" & trim(vMonth), 2) & "/" & _
            Right("0" & Trim(vDay), 2) & " " & _
            Right("0" & Trim(vHour), 2) & ":" & _
            Right("0" & Trim(vMin), 2)
End Function

'�ʐM�����̃Z�b�g
Function SetTusinDate()
	'�߂�l		[ O ]�ʐM����(������)
				'YYYYMMDDHHNNSS
	SetTusinDate = Trim(Year(Date)) & _
				   Trim(Right("0" & Month(Date), 2)) & _
				   Trim(Right("0" & Day(Date), 2)) & _
				   Trim(Right("0" & Hour(Time), 2)) & _
				   Trim(Right("0" & Minute(Time), 2)) & _
				   Trim(Right("0" & Second(Time), 2))

End Function
'���t������̎擾
Function GetYMDHM(DateVal)
	'DateVal	[ I ]���t�iDate�^�j
	'�߂�l		[ O ]���t������iYYYYMMDD�`���j

	GetYMDHM = _
		Year(DateVal) & _
		Right("0" & Month(DateVal), 2) & _
		Right("0" & Day(DateVal), 2)   & _
		Right("0" & Hour(DateVal), 2)  & _
		Right("0" & Minute(DateVal), 2)

End Function
'�w��N�����̃Z�b�g�Q
Function SetTusinDate2(sDate)
'	sDate		[ I ]���t(YYYY/MM/DD HH:MM)
	'�߂�l		[ O ]�w�����(YYYYMMDDHHMM)
				' --> 
	SetTusinDate2 = ""
	If sDate <> "" Then
		SetTusinDate2 = Left(Trim(sDate),4) & _
					   Mid(Trim(sDate), 6, 2) & _
					   Mid(Trim(sDate), 9, 2) & _
					   Mid(Trim(sDate), 12, 2) & _
					   Mid(Trim(sDate), 15, 2)
	End If
End Function

'���t�̕ϊ�����
Function ChangeDate(sDate)
'	sDate		[ I ]���t(YYYY/MM/DD/HH/MM etc)
	'�߂�l		[ O ]�w�����(YYYYMMDDHHMM)
				' --> 
	Dim sDat
	ChangeDate = ""
	If sDate = "////" Then
		Exit Function
	End If
response.write "sDate = "  & sDate
	sChangeWork =  Split(Trim(sDate),"/")

	ChangeDate = sChangeWork(0) & _
				 ArrangeNumV(sChangeWork(1),2) & _
				 ArrangeNumV(sChangeWork(2),2) & _
				 ArrangeNumV(sChangeWork(3),2) & _
				 ArrangeNumV(sChangeWork(4),2)

End Function

' ���l���w�肵�������̕�����ɕϊ�(�E�l�E�]���ɂ͂O)�R���e�i���͗p
Function ArrangeNumV(nNumber, nFigure)
	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNumV = sNum
End Function

%>