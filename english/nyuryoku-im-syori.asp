<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="vessel.inc"-->

<%
	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"
	' �g�����U�N�V�����h�c
	Const sTranID = "IM16"
	' �����敪
	Const sSyori = "R"
	' ���M�ꏊ
	Const sPlace = ""
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

	sSosin = Trim(Session.Contents("userid"))	'�C�݃R�[�h
    ' �G���[�t���O�̃N���A
    bError = false
    ' ���̓t���O�̃N���A
    bInput = true
    ' �w������̎擾
    Dim sContNo,sBLNo
    sContNo = UCase(Trim(Request.form("ContNo")))
    sBLNo = UCase(Trim(Request.form("BLNo")))
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	ConnectSvr conn, rsd
	if trim(sBLNo) = "" then
		sql = "SELECT ImportCont.VslCode, ImportCont.VoyCtrl, ImportCont.BLNo, VslSchedule.DsVoyage" & _
		      " FROM ImportCont, VslSchedule" & _
		      " WHERE ImportCont.ContNo = '" & sContNo & "'" & _
              " And VslSchedule.VslCode = ImportCont.VslCode" & _
		      " And VslSchedule.VoyCtrl = ImportCont.VoyCtrl"
				 
		'SQL�𔭍s���ėA���R���e�i������
		rsd.Open sql, conn, 0, 1, 1
		If Not rsd.EOF Then
		    sVslCode = Trim(rsd("VslCode"))		'�D��
		    sVoyCtrl = Trim(rsd("DsVoyage"))	'���q
		    sBLNo = Trim(rsd("BLNo"))			'BL�ԍ�
		Else
		    ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
		    bError = true
			strError = "�Y������R���e�i�����݂��܂���B"
		End If
	else
		sql = "SELECT BL.VslCode, BL.VoyCtrl, VslSchedule.DsVoyage" & _
		      " FROM BL, VslSchedule" & _
		      " WHERE BL.BLNo = '" & sBLNo & "'" & _
              " And VslSchedule.VslCode = BL.VslCode" & _
		      " And VslSchedule.VoyCtrl = BL.VoyCtrl"
				 
		'SQL�𔭍s����BL������
		rsd.Open sql, conn, 0, 1, 1
		If Not rsd.EOF Then
		    sVslCode = Trim(rsd("VslCode"))		'�D��
		    sVoyCtrl = Trim(rsd("DsVoyage"))	'���q
		Else
		    ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
		    bError = true
			strError = "�Y������BL No.�����݂��܂���B"
		End If
	end if
	rsd.Close
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/kaika3t.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
		<table>
			<tr> 
				<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
				<td nowrap><b>�R���e�i������</b></td>
				<td><img src="gif/hr.gif"></td>
			</tr>
		</table>
		<BR>
<%
	Dim sLogDate
	sLogDate = Trim(Request.form("Year")) & "/"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Month")),2) & "/"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Day")),2) & " "
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Hour")),2) & ":"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Min")),2)

    If bError Then
	    ' �G���[���b�Z�[�W�̕\��
	    DispErrorMessage strError 
	    strOption = sContNo & "/" & sBLNo & "/" & sLogDate & "," & "���͓��e�̐���:1(���)"

    Else
		'�g�����U�N�V�����t�@�C���쐬
	    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	    Dim sIM16, iSeqNo_IM16, strFileName, sTran, sTusin, sDate
		'�V�[�P���X�ԍ�
		iSeqNo_IM16 = GetDailyTransNo
		'�ʐM�����擾
		sTusin  = SetTusinDate
		sDate = Trim(Request.form("Year")) 
		sDate = sDate & Right("0" & Trim(Request.form("Month")),2)
		sDate = sDate & Right("0" & Trim(Request.form("Day")),2)
		sDate = sDate & Right("0" & Trim(Request.form("Hour")),2)
		sDate = sDate & Right("0" & Trim(Request.form("Min")),2)

		sIM16 = iSeqNo_IM16 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sContNo & "," & sBLNo & "," & sDate & ",," & sSosin
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM16
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sIM16
	    ti.Close
		Set ti = Nothing
	    ' �G���[���b�Z�[�W�̕\��
		strError = "����ɍX�V����܂����B"
        If strError="����ɍX�V����܂����B" Then
            DispInformationMessage strError
		    strOption = sContNo & "/" & sBLNo & "/" & sLogDate & "," & "���͓��e�̐���:0(������)"
        Else
            DispErrorMessage strError
		    strOption = sContNo & "/" & sBLNo & "/" & sLogDate & "," & "���͓��e�̐���:1(���)"
        End If

    End If

    ' �C�ݎ�����q�ɓ͂������w��
    WriteLog fs, "4004","�C�ݓ��͎�����q�ɓ�������","10", strOption
%>
			</center>
			<br>
		</td>
	</tr>
	<tr>
		<td valign="bottom"> 
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>
<%
%>