<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="vessel.inc"-->

<%
	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"
	' �g�����U�N�V�����h�c
	Const sTranID = "EX15"
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
    Dim sCntnrNo
    sCntnrNo = UCase(Trim(Request.form("CntnrNo")))
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ���̓R���e�iNo.�̃`�F�b�N
	ConnectSvr conn, rsd
	sql = "SELECT ExportCont.VslCode, ExportCont.VoyCtrl, VslSchedule.ShipLine, VslSchedule.LdVoyage" & _
	      " FROM ExportCont, VslSchedule" & _
	      " WHERE ExportCont.ContNo = '" & sCntnrNo & "' And VslSchedule.VslCode=ExportCont.VslCode And " & _
          "VslSchedule.VoyCtrl=ExportCont.VoyCtrl"
			 
	'SQL�𔭍s���ėA�o�R���e�i������
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
	    sVslCode = Trim(rsd("VslCode"))		'�D��
	    sVoyCtrl = Trim(rsd("LdVoyage"))	'���q
	Else
	    ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
	    bError = true
	    strError = "�Y������A�o�R���e�iNo.�����݂��܂���B"
	End If
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
					<td rowspan=2><img src="gif/kaika2t.gif" width="506" height="73"></td>
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
	sLogDate = Trim(Request.form("Year")) & "�N"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Month")),2) & "��"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Day")),2) & "��"

    If bError Then
	    ' �G���[���b�Z�[�W�̕\��
	    DispErrorMessage strError 
        strOption = sCntnrNo & "/" & sLogDate & "," & "���͓��e�̐���:1(���)"

    Else
		'�g�����U�N�V�����t�@�C���쐬
	    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	    Dim sEX15, iSeqNo_EX15, strFileName, sTran, sTusinsDate, sDate
		'�V�[�P���X�ԍ�
		iSeqNo_EX15 = GetDailyTransNo
		'�ʐM�����擾
		sTusin  = SetTusinDate
		sDate = Trim(Request.form("Year"))
		sDate = sDate & Right("0" & Trim(Request.form("Month")),2)
		sDate = sDate & Right("0" & Trim(Request.form("Day")),2) & "2359"

		sEX15 = iSeqNo_EX15 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sCntnrNo & "," & sDate
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX15
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sEX15
	    ti.Close
		Set ti = Nothing
	    ' �G���[���b�Z�[�W�̕\��
		strError = "����ɍX�V����܂����B"
        If strError="����ɍX�V����܂����B" Then
            DispInformationMessage strError
            strOption = sCntnrNo & "/" & sLogDate & "," & "���͓��e�̐���:0(������)"
        Else
            DispErrorMessage strError
            strOption = sCntnrNo & "/" & sLogDate & "," & "���͓��e�̐���:1(���)"
        End If

    End If
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
    ' �C��CY���������w��
    WriteLog fs, "4003","�C�ݓ���CY������","10", strOption
%>
