<%@ LANGUAGE="VBScript" %>
<%

' Added by Seiko-denki 2003.07.24
'�F�ؕ��@�ύX�̂��߃R�����g�A�E�g20040204 S
'	if request.querystring("UserId")<>"" then
'		strInputUserID = request.querystring("UserId")
'	else
'		if Session.Contents("userid") <> "" then
'			strInputUserID = Session.Contents("userid")
'		else
'			strInputUserID = ""
'		end if
'	end if
'
'	if strInputUserID<>"" and Session.Contents("login_count")=1 then
'		Session.Contents("userid")=strInputUserID
'	else
'		Session.Contents("login_count")=1
'		response.redirect "http://www.cont-info.com/Userchk2.asp?ReturnUrl=http://www.hits-h.com/SYWB013.asp"
'	end if
'�F�ؕ��@�ύX�̂��߃R�����g�A�E�g20040204 E
' End of Addition by Seiko-denki 2003.07.24


'Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB013.inc"-->
<!--#include file="SYWB077.inc"-->
<html>

<head>
<title>���o���\��\����ƈꗗ���</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
function SelDate(sel) {
//  location.href = "SYWB013.asp?TDATE=" + sel.options[sel.selectedIndex].value;
//	location.reload(true);
	location.replace("SYWB013.asp?UserId=<%=strInputUserID%>&TDATE=" + sel.options[sel.selectedIndex].value);
}
</SCRIPT>
</head>

<body>
<%
	'VP�Ή�2001/8/22
	'**********   �c�a�ڑ����   **********
	Dim conn, rsd			'�c�a�ڑ�
	'**********   ���t���   **********
	Dim sTrgDate			'�w����t("YYYYMMDD")
	Dim sDateNow			'���ݓ��t("YYYYMMDD")
	'**********   ���[�U���   **********
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator	'���[�U���
	'**********   �^�s�󋵏��   **********
	Dim iCurTime, iNextTime, iNextStat, iOpenSlot, sEndTime, iNextApp	'�^�s��
	'**********   �^�s���ԏ��   **********
	Dim sShtStart, sShtEnd	'�V���g���^�s���ԁiHHMM�j
	Dim iSTime, iETime		'�V���g���^�s���ԁi���ԑсj
	Dim iTimeCnt			'���ԑѐ�
	Dim TimeSlot(30)		'���ԑыL���i�C���f�b�N�X���s�ԍ��j��F08�`16,A,B,D
	Dim TimeNo(30)			'���ԑєԍ��i�C���f�b�N�X���s�ԍ��j��F8�`16,30,31,32
	Dim iRecDelCnt(30, 1)	'���o���{���i�C���f�b�N�X���s�ԍ��j
 	Dim sOpenFlag(23)		'�J���t���O�i�C���f�b�N�X�����ԑєԍ��j
	Dim iCloseMode(30)		'�������[�h�i�C���f�b�N�X���s�ԍ��j
							'�i0�F�^�s�O�@1�F�����@2�F�^�s���@3�F�m��@4�F�m�蒆�@-1�F�J�����j
	Dim TimeName(30), TimeJmp(30), sStatus(30)		'���ԑѕ\�����
	Dim iLuckChassis(30, 1), sLuckChassis(30, 1)	'�\����V���[�V��
	'**********   �f�[�^����   **********
	Dim iLineCnt(30)				'���ԑт��Ƃ̕\���s��
	Dim iRecIdx(30, 100)			'�\���s�Ɛ\�����̑Ή��e�[�u��
									'  0�`n�F�C���f�b�N�X�@-1�F�Ȃ��@-2�F�����Ȃ��@-3�F���o�Ȃ�
	'**********   ���̑�   **********
	Dim iEmptySlot, iEmptyChassis(1)			'�󂫃X���b�g�A��V���[�V
	Dim iCnt, iWk, sWk, bWk, i, k, sColor(9)
	Dim sDate
	Dim sDays(20), iDaysCnt						'�c�Ɠ�
	Dim sCell(16)
	'**********   �\�����i�C���f�b�N�X�����R�[�h�j   **********
	Dim iAppCnt					'�\�����
	Dim iAppOpeNo(1000)			'��Ɣԍ�
	Dim sAppUserNm(1000)		'���[�U��
	Dim sAppContNo(1000)		'�R���e�i�ԍ�
	Dim sAppBLNo(1000)			'�a�k�ԍ�
	Dim sAppRecDel(1000)		'���o���敪
	Dim sAppStatus(1000)		'���
	Dim sAppPlace(1000)			'�ꏊ
    Dim sAppChassisId(1000)		'�V���[�VID
	Dim sAppWorkFlag(1000)		'��ƒ��t���O
	Dim sAppCReason(1000)		'�L�����Z�����R
	Dim sAppTerm(1000)			'���ԑ�
	Dim sAppHopeTerm(1000)		'��]���ԑ�
	Dim iAppOpeOrder(1000)		'��Ə���
	Dim iAppDualOpeNo(1000)		'�f���A����Ɣԍ�
	Dim sAppContSize(1000)		'�R���e�i�T�C�Y
	Dim sAppFromTo(1000)		'���o��^������
	Dim sAppDelFlag(1000)		'�폜�t���O
	Dim sDelChaStock(1000)		'���o�w��V���[�V�̍݌�
	Dim sAppTerminal(1000)		'�^�[�~�i���R�[�h		'VP�Ή�(01/10/01)
	Dim sAppVPBookNo(1000)		'�u�o�u�b�L���O�ԍ�		'VP�Ή�(01/10/01)

	'�w����t�擾
	sTrgDate = TRIM(Request.QueryString("TDATE"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	If sGrpID = "" Then
		Response.Write "���[�U���o�^����Ă��܂���B(" & sUsrID & ")"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.End
	End If

	'�󂫃X���b�g�̎擾
	Call GetLackChassis(conn, rsd, sGrpID, _
			iEmptySlot, iEmptyChassis(0), iEmptyChassis(1))

	'���ݓ��t�擾
	sDateNow = GetYMDStr(Date())

	'�c�Ɠ��̎擾
	Call GetBusinessDays(conn, rsd, sDateNow, iDaysCnt, sDays)

	'�w����t���Ȃ��ꍇ�̓f�t�H���g���Z�b�g
	If sTrgDate = "" Then
		sTrgDate = Trim(sDays(1))
	End If

	'�^�s�󋵂��擾
	Call GetOpeStatusDtl(conn, rsd, _
						iCurTime, iNextTime, iNextStat, _
						iOpenSlot, sEndTime, iNextApp)

	'�[�ϏI���\����v�Z
	If sEndTime = "" Then
		sEndTime = "����"
	Else
		sEndTime = Left(sEndTime, 2) & ":" & Right(sEndTime, 2)
	End If

	If sTrgDate <> "WAIT" Then	'�ʏ�̏ꍇ
		'�O���[�v���ԑя��̎擾�i�J���̗L�����擾�j
		Call GetGrpSlot(conn, rsd, sGrpID, sTrgDate, sOpenFlag)

		'�V���g���^�s���Ԏ擾
		sShtStart = GetEnv(conn, rsd, "ShtStart")
		sShtEnd   = GetEnv(conn, rsd, "ShtEnd")
		iSTime = CLng(Left(sShtStart, 2))
		iETime = CLng(Left(sShtEnd, 2))
		If Right(sShtEnd, 2) = "00" Then
			iETime = iETime - 1
		End If

		'�V���g���^�s���ԑьv�Z
		iCnt = 0

		'���ԑт̌v�Z
		''�ߑO����
		For i = iSTime To 12
			TimeSlot(iCnt) = Right("0" & CStr(i), 2)
			TimeNo(iCnt) = i
			iCnt = iCnt + 1
		Next
		''�ߌ㎞��
		For i = 13 To iETime
			TimeSlot(iCnt) = Right("0" & CStr(i), 2)
			TimeNo(iCnt) = i
			iCnt = iCnt + 1
		Next
		''�ߌ�w��
		TimeSlot(iCnt) = "A"
		TimeNo(iCnt) = 30
		iCnt = iCnt + 1
		''�[�ώw��
		TimeSlot(iCnt) = "B"
		TimeNo(iCnt) = 31
		iCnt = iCnt + 1
		''���[�U�폜
		TimeSlot(iCnt) = "D"
		TimeNo(iCnt) = 32
		iCnt = iCnt + 1

		iTimeCnt = iCnt		'���ԑѐ�

		'�\�����擾
		iAppCnt = 0
		For i = 0 To iTimeCnt - 1
			'�w�莞�ԑсA�w��O���[�v�̐\�������擾
			Call GetAppHH(conn, rsd, _
					sGrpID, sTrgDate, TimeSlot(i), TimeNo(i), _
					sDateNow, iCurTime, iNextTime, iNextApp, _
					iRecDelCnt(i, 0), iRecDelCnt(i, 1), iCloseMode(i), _
					iAppCnt, _
					iAppOpeNo, sAppUserNm, sAppContNo, _
					sAppBLNo, sAppRecDel, sAppStatus, _
					sAppPlace, sAppChassisId, _
					sAppWorkFlag, sAppCReason, sAppContSize, _
					sAppTerm, sAppHopeTerm, iAppOpeOrder, _
					iAppDualOpeNo, sAppFromTo, sAppDelFlag, sDelChaStock, sAppTerminal, sAppVPBookNo)
		Next

		'�V���[�V�ݒ�
		''�f���A���Ŕ������V���[�V�����肵�Ă���ꍇ�ɔ��o���ɃV���[�V���Z�b�g
		Call SetAppChas(iAppCnt, _
						iAppOpeNo, sAppUserNm, sAppContNo, _
						sAppBLNo, sAppRecDel, sAppStatus, _
						sAppPlace, sAppChassisId, _
						sAppWorkFlag, sAppCReason, sAppContSize, _
						sAppTerm, sAppHopeTerm, iAppOpeOrder, _
						iAppDualOpeNo, sAppFromTo)

		'�{�����^�s���̏ꍇ�͉��V���[�V�v�Z���s��
		If sTrgDate = sDateNow Then
			'���V���[�V�v�Z
			Call CalcAppChas(conn, rsd, _
						sGrpID, sTrgDate, _
						iCurTime, iNextTime, iNextStat, _
						iAppCnt, _
						iAppOpeNo, sAppUserNm, sAppContNo, _
						sAppBLNo, sAppRecDel, sAppStatus, _
						sAppPlace, sAppChassisId, _
						sAppWorkFlag, sAppCReason, sAppContSize, _
						sAppTerm, sAppHopeTerm, iAppOpeOrder, _
						iAppDualOpeNo, sAppFromTo)
		End If

		'���ԑуZ���̐ݒ�
		For i = 0 To iTimeCnt - 1
			Call SetCell01( conn, rsd, sTrgDate, TimeSlot(i), _
							sShtStart, sShtEnd, iSTime, iETime, _
							iCloseMode(i), sOpenFlag, TimeName(i), TimeJmp(i), sStatus(i))
		Next
	Else		'�����҂��\���̏ꍇ
		'�w��O���[�v�̈����҂��\�����擾
		Call GetAppWait(conn, rsd, _
					sGrpID, _
					iAppCnt, _
					iAppOpeNo, sAppUserNm, sAppContNo, _
					sAppBLNo, sAppRecDel, sAppStatus, _
					sAppPlace, sAppChassisId, _
					sAppWorkFlag, sAppCReason, sAppContSize, _
					sAppTerm, sAppHopeTerm, iAppOpeOrder, _
					iAppDualOpeNo, sAppFromTo, sAppDelFlag, sAppTerminal, sAppVPBookNo)
	End If
%>

<IMG border=0 height=42 src="image/title01.gif" width=311>
<br><br>
<center>
<p><IMG border=0 height=34 src="image/title21.gif" width="236" height="34"><p>
</center>
<center>
	<b>��<% response.write sGrpName %>�O���[�v�̏��ł���<b>
</center><br>
<center>
          <TD align=middle height="36">
			<A href="SYWB017.asp?YMD=<%=sTrgDate%>">�V���g�����p��</A>�@�@�@
            <A href="../index.asp">���j���[��</A> 
          </TD>
</center><br>

<%
	rsd.Open "sUseDB", conn, 0, 1, 2
%>
<center>
<b>���݂̍݌ɏ��́@<U><%=Month(rsd("OutUpdtTime" & rsd("EnableDB")))%>��<%=Day(rsd("OutUpdtTime" & rsd("EnableDB")))%>���@
						<%=FormatDateTime(rsd("OutUpdtTime" & rsd("EnableDB")), vbShortTime)%></U>�@�̂��̂ł��@
�i<%=FormatDateTime(rsd("OutPUpdtTime"), vbShortTime)%>�ɍX�V�\��)�B
</b>

<%
	rsd.Close
%>
</center>

<%
	If sTrgDate <> "WAIT" Then	'�ʏ�\���̏ꍇ
		'�s�����Z�b�g
		For iCnt = 0 To iTimeCnt - 1		'���ԑѐ�
			iLineCnt(iCnt) = 0		'���ԑт��Ƃ̕\���s��
			bWk = False		'�f���A���t���O
			For i = 0 To iAppCnt - 1		'�\�����
				If sAppTerm(i) = TimeSlot(iCnt) Then	'���ԑт���v
					If TimeSlot(iCnt) <> "12" and _
					   TimeSlot(iCnt) <> "A" and _
					   TimeSlot(iCnt) <> "B" and _
					   TimeSlot(iCnt) <> "D" Then		'���Ԏw��̏ꍇ
						If sAppRecDel(i) = "R" Then			'�����̏ꍇ

							'�f�[�^�s�ǉ�
							iRecIdx(iCnt, iLineCnt(iCnt)) = i	'�\���s�Ɛ\�����̑Ή��e�[�u��
							iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'���ԑт��Ƃ̕\���s��

							If iAppDualOpeNo(i) > 0 Then		'�f���A���̏ꍇ
								bWk = True	'�f���A���t���O���I��
							ElseIf sAppStatus(i) <> "03" Then	'�P�ƂŃL�����Z���łȂ��ꍇ
								'��s�ǉ�
								iRecIdx(iCnt, iLineCnt(iCnt)) = -3	'�\���s�Ɛ\�����̑Ή��e�[�u��
								iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'���ԑт��Ƃ̕\���s��
							End If
						Else								'���o�̏ꍇ
							If bWk Then							'�f���A���̏ꍇ
								bWk = False		'�f���A���t���O���I�t
							ElseIf sAppStatus(i) <> "03" Then	'�P�ƂŃL�����Z���łȂ��ꍇ
								'��s�ǉ�
								iRecIdx(iCnt, iLineCnt(iCnt)) = -2	'�\���s�Ɛ\�����̑Ή��e�[�u��
								iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'���ԑт��Ƃ̕\���s��
							End If
							'�f�[�^�s�ǉ�
							iRecIdx(iCnt, iLineCnt(iCnt)) = i	'�\���s�Ɛ\�����̑Ή��e�[�u��
							iLineCnt(iCnt) = iLineCnt(iCnt) + 1	'���ԑт��Ƃ̕\���s��
						End If
					Else								'���Ԏw��ȊO�̏ꍇ
						iRecIdx(iCnt, iLineCnt(iCnt)) = i		'�\���s�Ɛ\�����̑Ή��e�[�u��
						iLineCnt(iCnt) = iLineCnt(iCnt) + 1		'���ԑт��Ƃ̕\���s��
					End If
				End If
			Next
			'���ԑт̕\���s�����O�̏ꍇ�͋�s��ǉ�
			If iLineCnt(iCnt) = 0 Then	'���ԑт��Ƃ̕\���s��
				iRecIdx(iCnt, iLineCnt(iCnt)) = -1	'�\���s�Ɛ\�����̑Ή��e�[�u��
				iLineCnt(iCnt) = 1					'���ԑт��Ƃ̕\���s��
			End If
		Next

		'�s���V���[�V���̌v�Z
		For iCnt = 0 To iTimeCnt - 1
			sLuckChassis(iCnt, 0) = "-"
			sLuckChassis(iCnt, 1) = "-"
			If iCloseMode(iCnt) <> 1 And _
			   TimeSlot(iCnt) <> "D" Then	'�����ȊO���폜�ȊO
				'��Ƃ̗L������
				bWk = False
				For i = 0 To iLineCnt(iCnt) - 1
					iWk = iRecIdx(iCnt, i)
					If iWk > -1 Then	'�󔒍s�łȂ�
						If sAppStatus(iWk) = "02" and _
						   sAppWorkFlag(iWk) <> "Y" and _
						   iAppDualOpeNo(iWk) = 0 Then

							bWk = True
							Exit For
						End If
					End If
				Next
				If bWk Then		'��Ƃ�����ꍇ�̂�
					'�󂫃V���[�V���̎擾working
					Call GetEmptyChassisCnt(conn, rsd, _
										sGrpID, _
										sTrgDate, _
										TimeSlot(iCnt), _
										iEmptyChassis(0), iEmptyChassis(1))
					sLuckChassis(iCnt, 0) = CStr(iEmptyChassis(0))
					sLuckChassis(iCnt, 1) = CStr(iEmptyChassis(1))
				End If
			End If
		Next
	Else		'�����҂��\���̏ꍇ
		'�s�����Z�b�g
		iTimeCnt = 1
		TimeSlot(0) = "X"
		TimeNo(0) = 0
		TimeName(0) = "�@"
		TimeJmp(0) = ""
		sStatus(0) = ""
		iCloseMode(0) = 1
		iRecDelCnt(0, 0) = "-"
		iRecDelCnt(0, 1) = "-"
		sLuckChassis(0, 0) = "-"
		sLuckChassis(0, 1) = "-"

		iLineCnt(0) = iAppCnt
		For i = 0 To iAppCnt - 1
			iRecIdx(0, i) = i
		Next
	End If

	'�c�a�ؒf
	conn.Close
%>
<br>
<center>
<table border="1">   

	<tr ALIGN=middle>
<td BGCOLOR=#F08080><select id=selectdate name=selectdate onChange="SelDate(this)">
<%
	'�c�Ɠ����j���[�쐬
	For iCnt = 0 To iDaysCnt - 1
		sWk = ""
		If sTrgDate = sDays(iCnt) Then
			sWk = "SELECTED"
		End If
%>
	<option <%=sWk%> VALUE ="<%=sDays(iCnt)%>"><%=ChgYMDStr3(sDays(iCnt))%></option>
<%
	Next

	sWk = ""
	If sTrgDate = "WAIT" Then
		sWk = "SELECTED"
	End If
%>
	<option <%=sWk%> VALUE = "WAIT">�����҂�</option>
</select></td>
		<td width="120" bgcolor ="#000080"><FONT COLOR="#ffffff">�J���c�g�i�{�j</FONT></td>
		<td width="50" BGCOLOR=#F08080><%=CStr(iOpenSlot)%></td>
		<td width="120" bgcolor ="#000080"><FONT COLOR="#ffffff">��X���b�g�i�{�j</FONT></td>
		<td width="50" BGCOLOR=#F08080><%=iEmptySlot%></td>
	</tr>
</table>
</center>
<br>
		<font face="�l�r �S�V�b�N">
		<center>
		<table border="1" width="930"  bgcolor = "#ffffff">   
			<tr ALIGN=middle bgcolor="#e8ffe8">
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>��Ǝ��ԁ@</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>�S�{��<br>��/�o</TH>
			    <TH BGCOLOR=#7FFFD4 COLSPAN=2>�󼬰�<br>�ߕs��</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>�\��<br>�ԍ�</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2 width="20">����</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>�R���e�i�^�a�k<br>�^�u�b�L���O</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>�\��<br>�^�C�v</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2 width="20">���</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2 width="20">�T�C�Y</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2><A href="SYWB021.asp?YMD=<%=sTrgDate%>">����ID</A></TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>�Ώ�<br>�b�x�^�u�o</TH>		<!--��o���Ή� -->
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>�ꏊ</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>���</TH>
			    <TH BGCOLOR=#7FFFD4 ROWSPAN=2>���l</TH>
			</tr>
			<tr ALIGN=middle bgcolor="#e8ffe8">
			    <TH BGCOLOR=#7FFFD4>20</TH>
			    <TH BGCOLOR=#7FFFD4>40</TH>
			</tr>
<%
	For iCnt = 0 To iTimeCnt - 1	'���ԑѐ�
		iWk = iLineCnt(iCnt)			'���ԑт��Ƃ̕\���s��
		If iWk > 0 Then
			'�f�[�^�Z���̐ݒ�(01/10/02 VP�Ή�)
			Call SetCell05(iRecIdx(iCnt, 0), iCloseMode(iCnt), _
					sTrgDate, TimeSlot(iCnt), TimeName(iCnt), _
					iAppOpeNo,  sAppUserNm, sAppContNo, sAppBLNo, _
          			sAppRecDel, sAppStatus,  sAppPlace, _
          			sAppChassisId, sAppWorkFlag, sAppCReason, _
          			sAppContSize, sAppTerm, sAppHopeTerm, _
					iAppOpeOrder, iAppDualOpeNo, sAppFromTo, _
					sAppDelFlag, sDelChaStock, sAppTerminal, sAppVPBookNo, _
					sCell)
			'���ԑуZ���J���[�̌v�Z
			sColor(0) = ""
			sColor(1) = ""
			sColor(3) = "bgcolor=""#AFEEEE"" "
			If TimeSlot(iCnt) = "D" Then
				sColor(0) = "bgcolor=""#dda0dd"" "
				sColor(3) = "bgcolor=""#dda0dd"" "
			End If
			If TimeSlot(iCnt) = "12" Or _
			   TimeSlot(iCnt) = "A" Or _
			   TimeSlot(iCnt) = "B" Then
				sColor(1) = "bgcolor=""#FFD700"" "
			Else
				If TimeSlot(iCnt) = "D" Then
					sColor(1) = "bgcolor=""#dda0dd"" "
				Else
					If sStatus(iCnt) = "����" Then
						sColor(1) = "bgcolor=""#c0c0c0"" "
					ElseIf sStatus(iCnt) = "�^�s��" Then
						sColor(1) = "bgcolor=""#F08080"" "
					Else
						sColor(1) = "bgcolor=""#FFFFE0"" "
					End If
				End If
			End If
			'�f�[�^�Z���J���[�̌v�Z
			Call CalcDataColor(sColor(2), sCell)
%>
			<tr ALIGN=middle <%=sColor(0)%>>
<% '���[�U�폜�Z���^�\��(2001/03/23)
				If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
			    <td <%=sColor(1)%> ROWSPAN=2>
							<%=TimeJmp(iCnt) & TimeName(iCnt)%></A>
			    <td <%=sColor(3)%> ROWSPAN=2><%=iRecDelCnt(iCnt, 0)%>/<%=iRecDelCnt(iCnt, 1)%><br></td>
<%				Else	
					If TimeName(iCnt) =  "���[�U�폜" Then	%>
				    <td <%=sColor(1)%> ROWSPAN=<%=iWk%>>
								<%=TimeJmp(iCnt) & TimeName(iCnt)%></A>
						<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=iRecDelCnt(iCnt, 0)%>/<%=iRecDelCnt(iCnt, 1)%></td>
<%					Else	%>
						<td <%=sColor(1)%> ROWSPAN=<%=iWk%>>
									<%=TimeJmp(iCnt) & TimeName(iCnt)%></A><br>
									<%=sStatus(iCnt)%></A></td>
						<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=iRecDelCnt(iCnt, 0)%>/<%=iRecDelCnt(iCnt, 1)%></td>
<%					End If	%>
<%				End If	%>
<% '��V���[�V�ߕs���}�C�i�X���Ԏ��\��(2001/03/09)
				If sLuckChassis(iCnt, 0) <> "-" Then
					If sLuckChassis(iCnt, 0) < 0 Then	
						If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
					    <td <%=sColor(3)%>  ROWSPAN=2><FONT color=Red><B><%=sLuckChassis(iCnt, 0)%></B></FONT></td>
<%						Else	%>
					    <td <%=sColor(3)%>  ROWSPAN=<%=iWk%>><FONT color=Red><B><%=sLuckChassis(iCnt, 0)%></B></FONT></td>
<%						End if	
					Else				
						If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
					    <td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 0)%></td>
<%						Else	%>
					    <td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 0)%></td>
<%						End if
					End If					
				Else						
					If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
					<td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 0)%><br></td>
<%					Else	%>
					<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 0)%></td>
<%					End If
				End If					

				If sLuckChassis(iCnt, 1) <> "-" Then
					If sLuckChassis(iCnt, 1) < 0 Then	
						If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
					    <td <%=sColor(3)%>  ROWSPAN=2><FONT color=Red><B><%=sLuckChassis(iCnt, 1)%></B></FONT></td>
<%						Else	%>
					    <td <%=sColor(3)%>  ROWSPAN=<%=iWk%>><FONT color=Red><B><%=sLuckChassis(iCnt, 1)%></B></FONT></td>
<%						End if	
					Else				
						If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
					    <td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 1)%></td>
<%						Else	%>
					    <td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 1)%></td>
<%						End if
					End If					
				Else						
					If TimeName(iCnt) =  "���[�U�폜" And iWk <  2 Then	%>
					<td <%=sColor(3)%> ROWSPAN=2><%=sLuckChassis(iCnt, 1)%><br></td>
<%					Else	%>
					<td <%=sColor(3)%> ROWSPAN=<%=iWk%>><%=sLuckChassis(iCnt, 1)%></td>
<%					End If
				End If	%>				

			    <td <%=sColor(2)%>><%=sCell(1)%></td>
			    <td <%=sColor(2)%>><%=sCell(2)%></td>
			    <td <%=sColor(2)%>><%=sCell(3)%></td>
			    <td <%=sColor(2)%>><%=sCell(4)%></td>
			    <td <%=sColor(2)%>><%=sCell(5)%></td>
			    <td <%=sColor(2)%>><%=sCell(6)%></td>
			    <td <%=sColor(2)%>><%=sCell(7)%></td>
			    <td <%=sColor(2)%>><%=sCell(0)%></td>			<!--��o���Ή� -->
			    <td <%=sColor(2)%>><%=sCell(8)%></td>
			    <td <%=sColor(2)%>><%=sCell(9)%></td>
			    <td <%=sColor(2)%>><%=sCell(10)%></td>
			</tr>
<%
		End If
		For i = 1 To iWk - 1		'���ԑт��Ƃ̕\���s��-1
			'�f�[�^�Z���̐ݒ�
			Call SetCell05(iRecIdx(iCnt, i), iCloseMode(iCnt), _
					sTrgDate, TimeSlot(iCnt), TimeName(iCnt), _
					iAppOpeNo, sAppUserNm, sAppContNo, sAppBLNo, _
          			sAppRecDel, sAppStatus, sAppPlace, _
          			sAppChassisId, sAppWorkFlag, sAppCReason, _
          			sAppContSize, sAppTerm, sAppHopeTerm, _
					iAppOpeOrder, iAppDualOpeNo, sAppFromTo, _
					sAppDelFlag, sDelChaStock, sAppTerminal, sAppVPBookNo, _
					sCell)
			'�f�[�^�Z���J���[�̌v�Z
			Call CalcDataColor(sColor(2), sCell)
%>

			<tr ALIGN=middle <%=sColor(0)%>>
			    <td <%=sColor(2)%>><%=sCell(1)%></td>
			    <td <%=sColor(2)%>><%=sCell(2)%></td>
			    <td <%=sColor(2)%>><%=sCell(3)%></td>
			    <td <%=sColor(2)%>><%=sCell(4)%></td>
			    <td <%=sColor(2)%>><%=sCell(5)%></td>
			    <td <%=sColor(2)%>><%=sCell(6)%></td>
			    <td <%=sColor(2)%>><%=sCell(7)%></td>
			    <td <%=sColor(2)%>><%=sCell(0)%></td>			<!--��o���Ή� -->
			    <td <%=sColor(2)%>><%=sCell(8)%></td>
			    <td <%=sColor(2)%>><%=sCell(9)%></td>
			    <td <%=sColor(2)%>><%=sCell(10)%></td>
			</tr>
<%
		Next
	Next
%>
		</table>
		</center>
		</font>

<br>     
<br>     
</body>     
</html>     
