<% 
'�A�o�R���e�i�Ɖ�
Function SerchKCntnr(conn, rsd, pfp, sCont, sSeal, sJyuryo, sSoJyuryo)
    'conn       [ I ]DB�ڑ�Object
    'rsd        [ I ]���R�[�h�Z�b�gObject
    'pfp        [ I ]�e���|�����t�@�C��
    'sCont     	[ I ]�R���e�iNo.
    'sSeal     	[ I ]�V�[��No.
    'sJyuryo    [ I ]�d��

    bWrite = 0        '�o�̓��R�[�h����

    ' ���̓R�[���T�C���̃`�F�b�N
	ConnectSvr conn, rsd
	sql = "SELECT ExportCont.VslCode, ExportCont.VoyCtrl, ExportCont.BookNo, Container.ShipLine, ExportCont.WHArTime, "
	sql = sql & " VslSchedule.LdVoyage "
	sql = sql & " FROM ExportCont, Container, VslSchedule"
	sql = sql & " WHERE ExportCont.ContNo='" & sCont & "' And Container.VslCode = ExportCont.VslCode"
	sql = sql & " AND Container.VoyCtrl = ExportCont.VoyCtrl"
	sql = sql & " AND Container.ContNo = ExportCont.ContNo"
	sql = sql & " AND Container.VslCode = VslSchedule.VslCode"
	sql = sql & " AND Container.VoyCtrl = VslSchedule.VoyCtrl"

	'SQL�𔭍s���ėA�o�R���e�i������
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
		strOut = Trim(rsd("VslCode")) & ","						' 1:�D��
        strOut = strOut & Trim(rsd("LdVoyage")) & ","			' 2:�ςݎ��q�ԍ�
        strOut = strOut & Trim(sCont) & ","            			' 3:�R���e�iNo.
        strOut = strOut & Trim(rsd("BookNo")) & ","             ' 4:Booking No.
		strOut = strOut & Trim(rsd("ShipLine")) & ","			' 5:�D��
		strOut = strOut & GetYMDHM(rsd("WHArTime")) & ","		' 6:�o���l�ߓ���
        strOut = strOut & sJyuryo & ","         				' 7:�d��
        strOut = strOut & sSeal & ","	        				' 8:�V�[��No.
        strOut = strOut & sSoJyuryo   		      				' 9:���d��

        pfp.WriteLine strOut
        bWrite = bWrite + 1

	End If
	rsd.Close

    SerchKCntnr = bWrite    '�Y�����R�[�h���������ǂ���
End Function

'�{�D���ÁE�{�D��`�n�Ɖ�
Function SerchKVsl(conn, rsd, pfp, sVsl, sVoy, sPort, sEta, sTa, sTD, sTaL, sTdL)
    'conn       [ I ]DB�ڑ�Object
    'rsd        [ I ]���R�[�h�Z�b�gObject
    'pfp        [ I ]�e���|�����t�@�C��
    'sVsl     	[ I ]�D���i�R�[���T�C���j
    'sVoy     	[ I ]�g���`�^�ςݍ`
    'sPort    	[ I ]��`�n
    'sEta     	[ I ]ETA
    'sTa    	[ I ]TA
    'sTD     	[ I ]TD
    'sTaL    	[ I ]�ݽ����`�\���
    'sTdL    	[ I ]�ݽ��o�`�\���

    bWrite = 0        '�o�̓��R�[�h����

    ' ���̓R�[���T�C���̃`�F�b�N
	ConnectSvr conn, rsd
	sql = "SELECT VslSchedule.VslCode, VslSchedule.VoyCtrl, VslSchedule.DsVoyage, VslSchedule.LdVoyage, VslSchedule.ShipLine, "
	sql = sql & " VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.TD, VslPort.ETALong, VslPort.ETDLong"
	sql = sql & " FROM VslSchedule, VslPort"
	sql = sql & " WHERE ( VslSchedule.VslCode='" & sVsl & "' And VslSchedule.DsVoyage = '" & sVoy & "'"
	sql = sql & " AND VslPort.PortCode = '" & sPort & "'"
	sql = sql & " AND VslSchedule.VslCode = VslPort.VslCode"
	sql = sql & " AND VslSchedule.VoyCtrl = VslPort.VoyCtrl ) "
	sql = sql & " Or ( VslSchedule.VslCode='" & sVsl & "' And VslSchedule.LdVoyage = '" & sVoy & "'"
	sql = sql & " AND VslPort.PortCode = '" & sPort & "'"
	sql = sql & " AND VslSchedule.VslCode = VslPort.VslCode"
	sql = sql & " AND VslSchedule.VoyCtrl = VslPort.VoyCtrl ) "

	'SQL�𔭍s���ėA�o�R���e�i������
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
		strOut = Trim(sVsl) & ","								' 1:�D��
        strOut = strOut & Trim(rsd("DsVoyage")) & ","			' 2:�g�����q�ԍ�
        strOut = strOut & Trim(rsd("LdVoyage")) & ","			' 3:�ςݎ��q�ԍ�
        strOut = strOut & Trim(rsd("ShipLine")) & ","			' 4:�D��
        strOut = strOut & Trim(sPort) & ","            			' 5:��`�n��
        strOut = strOut & ChangeDate(sTaL,8) & ","         		' 6:����Long Schedule
        strOut = strOut & ChangeDate(sTdL,8) & "," 		        ' 7:����Long Schedule
        strOut = strOut & ChangeDate(sEta,12) & ","             	' 8:���ݗ\�����
		strOut = strOut & ChangeDate(sTa,12)  & ","				' 9:���݊�������
		strOut = strOut & ChangeDate(sTD,12)						' 10:���݊�������

        pfp.WriteLine strOut
        bWrite = bWrite + 1

	End If
	rsd.Close

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
Function ChangeDate(sDate,ilength)
'	sDate		[ I ]���t(YYYY/MM/DD/HH/MM etc)
'	ilength		[ I ]�����@�W���܂��͂P�Q��
'	�߂�l		[ O ]�w�����(YYYYMMDDHHMM)

    Dim sDat
    ChangeDate = ""
    If sDate = "////" Or sDate="" Then
        Exit Function
    End If
    sChangeWork =  Split(Trim(sDate),"/")

    If (ilength = 12 And UBound(sChangeWork)<>4) Or (ilength = 8 And UBound(sChangeWork)<>2) Then
        ChangeDate = "�t�H�[�}�b�g���ُ�ł��B(" & sDate & ")"
        Exit Function
    End If

    iYear = CInt(sChangeWork(0))
    If iYear<1990 Or iYear>2100 Then
        ChangeDate = "�N���ُ�ł��B(" & sChangeWork(0) & ")"
        Exit Function
    End If
    iMonth = CInt(sChangeWork(1))
    If iMonth<1 Or iMonth>12 Then
        ChangeDate = "�����ُ�ł��B(" & sChangeWork(1) & ")"
        Exit Function
    End If
    iDay = CInt(sChangeWork(2))
    If iDay<1 Or iDay>31 Then
        ChangeDate = "�����ُ�ł��B(" & sChangeWork(2) & ")"
        Exit Function
    End If
    If iMonth=4 Or iMonth=6 Or iMonth=9 Or iMonth=11 Then
        If iDay>30 Then
            ChangeDate = "�����ُ�ł��B(" & sChangeWork(2) & ")"
            Exit Function
        End If
    End If
    If iMonth=2 Then
        If ( (iYear Mod 4)=0 And (iYear Mod 100)<>0 ) Or (iYear Mod 400)=0 Then
            If iDay>29 Then
                ChangeDate = "�����ُ�ł��B(" & sChangeWork(2) & ")"
                Exit Function
            End If
        Else
            If iDay>28 Then
                ChangeDate = "�����ُ�ł��B(" & sChangeWork(2) & ")"
                Exit Function
            End If
        End If
    End If
    If ilength>8 Then
        If CInt(sChangeWork(3))<0 Or CInt(sChangeWork(3))>23 Then
            ChangeDate = "�����ُ�ł��B(" & sChangeWork(3) & ")"
            Exit Function
        End If
        If CInt(sChangeWork(4))<0 Or CInt(sChangeWork(4))>59 Then
            ChangeDate = "�����ُ�ł��B(" & sChangeWork(4) & ")"
            Exit Function
        End If
    End If
	If ilength = 12 Then
		ChangeDate = ArrangeNumV(iYear,4) & _
					 ArrangeNumV(iMonth,2) & _
					 ArrangeNumV(iDay,2) & _
					 ArrangeNumV(sChangeWork(3),2) & _
					 ArrangeNumV(sChangeWork(4),2)
	Else
		ChangeDate = ArrangeNumV(iYear,4) & _
					 ArrangeNumV(iMonth,2) & _
					 ArrangeNumV(iDay,2)
	End If
End Function

' ���l���w�肵�������̕�����ɕϊ�(�E�l�E�]���ɂ͂O)�R���e�i���͗p
Function ArrangeNumV(nNumber, nFigure)
	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNumV = sNum
End Function

%>