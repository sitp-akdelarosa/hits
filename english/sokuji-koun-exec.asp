<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	�������o�V�X�e���y�`�^�p�z	�ύX,�폜�p���

%>

<%
	' �Z�b�V�����̃`�F�b�N
	CheckLogin "sokuji.asp"

	' �`�^�R�[�h�擾
	sOpe = Trim(Session.Contents("userid"))

	Dim sSend,sStop,iSend,iChkCount
	sSend = Trim(Request.form("send"))
	sStop = Trim(Request.form("stop"))

	If sStop<>"" Then
		Response.Redirect "sokuji-koun-list.asp"
		Response.End
	Else
		' �g�����U�N�V�����t�@�C���̊g���q 
		Const SEND_EXTENT = "snd"
		' �g�����U�N�V�����h�c
		Const sTranID = "IM20"
		' ���M�ꏊ
		Const sPlace = ""
		' �Z�b�V�����̃`�F�b�N
		iSend = Trim(Session.Contents("send"))
		iChkCount=Session.Contents("ChkCount")

		' �G���[�t���O�̃N���A
		bError = false
		' ���̓t���O�̃N���A
		bInput = true

		' File System Object �̐���
		Set fs=Server.CreateObject("Scripting.FileSystemobject")

		' �w������̎擾
		Dim sShipper(),sShipLine(),sVslCode(),sBL(),sCont(),sForwarder(),sLineNo()
		For i=1 to iChkCount
			ReDim Preserve sShipper(i)
			ReDim Preserve sShipLine(i)
			ReDim Preserve sVslCode(i)
			ReDim Preserve sBL(i)
			ReDim Preserve sCont(i)
			ReDim Preserve sForwarder(i)
			ReDim Preserve sLineNo(i)
			sShipper(i) 	= UCase(Trim(Session.Contents("shipper" & i)))
			sShipLine(i) = UCase(Trim(Session.Contents("shipline" & i)))
			sVslCode(i) = UCase(Trim(Session.Contents("vslcode" & i)))
			sBL(i) = UCase(Trim(Session.Contents("bl" & i)))
			sCont(i) = UCase(Trim(Session.Contents("cont" & i)))
			sForwarder(i)	= UCase(Trim(Session.Contents("forwarder" & i)))
			sLineNo(i)	= UCase(Trim(Session.Contents("lineno" & i)))
		Next

		Dim sReject,sRecschTime,sYear,sMonth,sDay,sHour,sMin
		sYear	= GetNumStr(UCase(Trim(Request.form("year"))),4)
		sMonth	= GetNumStr(UCase(Trim(Request.form("month"))),2)
		sDay	= GetNumStr(UCase(Trim(Request.form("day"))),2)
		sHour	= GetNumStr(UCase(Trim(Request.form("hour"))),2)
		sMin	= GetNumStr(UCase(Trim(Request.form("min"))),2)
		If iSend=0 Then
			sReject="0"
			sRecschTime=sYear & sMonth & sDay & sHour & sMin
		Else
			sReject="1"
			sRecschTime=""
		End If

'		' ���p�J���}�`�F�b�N
'		If InStr(sShipper,",")<>0 Or InStr(sShipLine,",")<>0 Or InStr(sVslCode,",")<>0 Or _
'			InStr(sBL,",")<>0 Or InStr(sCont,",")<>0 _
'		Then
'			bError = true
'			strError = "���͂̍ہA���p�J���}�͎g�p���Ȃ��ŉ������B"
'		Else

' �g�����U�N�V�����t�@�C���쐬
			' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
			Dim sIM20, iSeqNo_IM20, strFileName, sTran, sTusin, sDate, iSeqCnt
			'�V�[�P���X�ԍ�
			iSeqNo_IM20 = GetDailyTransNo
			'�ʐM�����擾
			sTusin  = SetTusinDate
			' �����敪
			sSyori="R"
			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM20
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
			Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
			iSeqCnt=0
			For i=1 to iChkCount
				If sLineNo(i)<>"" Then
					' 2�s����̓V�[�P���X�ԍ���ύX
					If iSeqCnt<>0 Then iSeqNo_IM20 = GetDailyTransNo
					sIM20 = iSeqNo_IM20 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & sOpe & ",," & _
					sShipper(i) & "," & sShipLine(i) & "," & sVslCode(i) & "," &  sBL(i) & "," & sCont(i) & "," & _
					sReject & "," & sRecschTime' & "," & sForwarder(i)
					ti.WriteLine sIM20
					iSeqCnt=iSeqCnt+1
				End If
			Next
			ti.Close
			Set ti = Nothing
' �g�����U�N�V���������܂�


' Temp�t�@�C���쐬

'			For i=1 to iChkCount
'				If sBL(i)="" Then sBL(i)="*"
'				If sCont(i)="" Then sCont(i)="*"
'			Next

			' File System Object �̐���
			Set fs=Server.CreateObject("Scripting.FileSystemobject")

			Dim strTempFileName
			' �\���t�@�C���̎擾
			strTempFileName = Session.Contents("tempfile")
			If strTempFileName="" Then
				' �Z�b�V�������؂�Ă���Ƃ�
				Response.Redirect "sokuji-koun-updtchk.asp"             '���j���[��ʂ�
				Response.End
			End If

			strTempFileName="./temp/" & strTempFileName

			' �\���t�@�C����Open
			Set ti=fs.OpenTextFile(Server.MapPath(strTempFileName),1,True)

			' �ڍו\���s�̃f�[�^�̎擾
			Dim strData(),LineNo
			LineNo=0
			Do While Not ti.AtEndOfStream
				strTemp=ti.ReadLine
				ReDim Preserve strData(LineNo)
				strData(LineNo) = strTemp
				LineNo=LineNo+1
			Loop
			ti.Close

			'' �}�X�^DB����̓ǂݍ���
			ConnectSvr conn, rsd
			For i=1 to iChkCount
				ReDim Preserve sShipperAbrev(i)
				ReDim Preserve sShipLineAbrev(i)
				ReDim Preserve sVesselAbrev(i)
				'' DB�̓ǂݍ���
				sql = "SELECT NameAbrev FROM mShipper WHERE Shipper='" & sShipper(i) & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sShipperAbrev(i) = Trim(rsd(0))
				  rsd.MoveNext
				Loop
				rsd.Close
				'' DB�̓ǂݍ���
				sql = "SELECT NameAbrev FROM mShipLine WHERE ShipLine='" & sShipLine(i) & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sShipLineAbrev(i) = Trim(rsd(0))
				  rsd.MoveNext
				Loop
				rsd.Close
				'' DB�̓ǂݍ���
				sql = "SELECT FullName FROM mVessel WHERE VslCode='" & sVslCode(i) & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sVesselAbrev(i) = Trim(rsd(0))
				  rsd.MoveNext
				Loop
				rsd.Close
			Next

			'' �t�@�C����������
			ReDim anyTmp(10)
			Set ti=fs.OpenTextFile(Server.MapPath(strTempFileName),2,True)
			For i=1 To LineNo
				If sLineNo(i)<>"" Then
'					anyTmp=Split(strData(i-1),",")
					anyTmp(0) = sShipperAbrev(i)
					anyTmp(1) = sShipLineAbrev(i)
					anyTmp(2) = sVesselAbrev(i)
					anyTmp(3) = sBL(i)
					anyTmp(4) = sCont(i)
					If sReject="0" Then
						anyTmp(5) = "��"
					ElseIf sReject="1" Then
						anyTmp(5) = "�~"
					Else
						anyTmp(5) = ""
					End If
					If sRecschTime<>"" Then
						anyTmp(6) = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin
					Else
						anyTmp(6) = ""
					End If
					anyTmp(7) = sShipper(i)
					anyTmp(8) = sShipLine(i)
					anyTmp(9) = sVslCode(i)
					anyTmp(10) = sForwarder(i)
					strTemp=anyTmp(0)
					For j=1 To UBound(anyTmp)
					    strTemp=strTemp & "," & anyTmp(j)
					Next
					ti.WriteLine strTemp
				Else
				  ti.WriteLine strData(i-1)
				End If
			Next
			ti.Close

'		End If
	End If

	' Log�쐬
	If iSend=0 Then
		WriteLog fs, "7004", "�������o�V�X�e��-�`�^�p�\�莞������", "10", sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ","
	Else
		WriteLog fs, "7003", "�������o�V�X�e��-�`�^�p���ꗗ", "10", ","
	End If

	Response.Redirect "sokuji-koun-list.asp"
	Response.End

%>
