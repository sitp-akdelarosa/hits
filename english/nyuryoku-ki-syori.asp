<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->

<%
	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"
	' �g�����U�N�V�����h�c
	Const sTranID = "EX05"
	' �����敪
	Const sSyori = "R"
	' ���M�ꏊ
	Const sPlace = ""

    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-ki.asp"
	sSosin = Trim(Session.Contents("userid"))

    ' �G���[�t���O�̃N���A
    bError = false

    ' ���̓t���O�̃N���A
    bInput = true

    ' �w������̎擾
    Dim sContNo, sSealNo, sJyuryo

    sContNo = UCase(Trim(Request.form("ContNo")))
    sSealNo = UCase(Trim(Request.form("SealNo")))
	If InStr(sSealNo,",")<>0 Then
	    bError = true
	    strError = "�V�[��No.�ɂ͔��p�J���}�����Ȃ��ŉ������B"
	End If

    sJyuryo = UCase(Trim(Request.form("Jyuryo")))
    If sJyuryo<>"" Then
        sJyuryo=CInt(CDbl(sJyuryo)*10)
    End If
    sSoJyuryo = UCase(Trim(Request.form("SoJyuryo")))
    If sSoJyuryo<>"" Then
        sSoJyuryo=CInt(CDbl(sSoJyuryo)*10)
    End If
    sRefer = UCase(Trim(Request.form("rf")))
    sDG = UCase(Trim(Request.form("dg")))
' 2/9�ύX�������ǉ�
	If sRefer="ON" Then
		sRefer = 1
	Else
		sRefer = 2
	End If
	If sDG="ON" Then
		sDG =1
	Else
		sDG = 2
	End If
' �����܂�
    sRefDB=""
    If sRefer="1" And sDG="1" Then
        sRefDB="RH"
    ElseIf sRefer="1" Then
        sRefDB="R"
    ElseIf sDG="1" Then
        sRefDB="H"
    End If

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ���̓R���e�iNo.�̃`�F�b�N
	ConnectSvr conn, rsd
	sql = "SELECT ExportCont.VslCode, ExportCont.VoyCtrl, ExportCont.BookNo, ExportCont.WHArTime, VslSchedule.LdVoyage, VslSchedule.ShipLine "
	sql = sql & " FROM ExportCont, VslSchedule"
	sql = sql & " WHERE ExportCont.ContNo='" & sContNo & "' And VslSchedule.VslCode = ExportCont.VslCode"
	sql = sql & " AND VslSchedule.VoyCtrl = ExportCont.VoyCtrl"
			 
	'SQL�𔭍s���ėA�o�R���e�i������
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
	    sVslCode = Trim(rsd("VslCode"))		'�D��
	    sVoyCtrl = Trim(rsd("LdVoyage"))	'���q
	    sBookNo = Trim(rsd("BookNo"))		'�u�b�L���O
	    stShipLine = Trim(rsd("ShipLine"))	'�D��
'	    stWHArTime = GetYMDHM(rsd("WHArTime")) 		'�o���l�ߓ���
	    strOption = sContNo & "/" & sSealNo & "/" & sJyuryo & "/" & sSoJyuryo & "/" & sRefer & "/" & sDG & "," & "���͓��e�̐���:0(������)"
	Else
	    ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
	    bError = true
	    strError = "�Y������R���e�i�����݂��܂���B"
	    strOption = sContNo & "/" & sSealNo & "/" & sJyuryo & "/" & sSoJyuryo & "/" & sRefer & "/" & sDG & "," & "���͓��e�̐���:1(���)"
	End If
	rsd.Close

    ' �C�ݓ��̓V�[��No.�A�d�ʓ���
    WriteLog fs, "4002","�C�ݓ��̓V�[��No.�E�d�ʓ���","10", strOption

    If Not bError Then
		'�g�����U�N�V�����t�@�C���쐬

	    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	    Dim sEX05, iSeqNo_EX05, strFileName, sTran, sTusin

		'�V�[�P���X�ԍ�
		iSeqNo_EX05 = GetDailyTransNo
		'�ʐM�����擾
		sTusin  = SetTusinDate

		sEX05 = ""
		sEX05 = iSeqNo_EX05 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sContNo & "," & sBookNo & "," & stShipLine & "," & stWHArTime & "," & _
				sSoJyuryo & "," & sSealNo & "," & sJyuryo & "," & _
                sSosin & ",," & sRefDB

		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX05
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		strFileName_01=Server.MapPath(strFileName_01)
	    Set ti=fs.OpenTextFile(strFileName_01,2,True)
		ti.WriteLine sEX05
	    ti.Close
		Set ti = Nothing

	    ' �G���[���b�Z�[�W�̕\��
		strError = "����ɍX�V����܂����B"
	End If
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>

<!--
function gotoURL(){
    var gotoUrl=document.con.select.options[document.con.select.selectedIndex].value
    document.location.href=gotoUrl 
}
//-->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
	<td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika1t.gif" width="506" height="73"></td>
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
    If bError Then
	    ' �G���[���b�Z�[�W�̕\��
	    DispErrorMessage strError

    Else
        If strError="����ɍX�V����܂����B" Then
            DispInformationMessage strError
        Else
            DispErrorMessage strError
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
