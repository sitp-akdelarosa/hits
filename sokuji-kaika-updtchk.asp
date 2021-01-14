<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	�������o�V�X�e���y�C�ݗp�z	�X�V���\���`�F�b�N�ATemp�t�@�C���쐬

%>

<%
	' �Z�b�V�����̃`�F�b�N
	CheckLogin "sokuji.asp"

	' �C�݃R�[�h�擾
	sForwarder = Trim(Session.Contents("userid"))

	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

  ' �\���t�@�C���̎擾
	Dim strFileName
  'strFileName = Session.Contents("tempfile")

	' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
	Session.Contents("tempfile")=strFileName

	ConnectSvr conn, rsd

	Dim QdelNo
	LineNo = 0

	'' DB�̓ǂݍ���
	sql = "SELECT mShipper.NameAbrev,mShipLine.NameAbrev,mVessel.FullName," & _
				"QuickDel.BLNo,QuickDel.ContNo,QuickDel.RejectFlag,QuickDel.RecSchTime," & _
				"QuickDel.Shipper,QuickDel.ShipLine,QuickDel.VslCode,BL.OpeCode " & _
				"FROM QuickDel,mShipLine,mVessel,mShipper,BL " & _
				"WHERE mShipLine.ShipLine=*QuickDel.ShipLine AND mVessel.VslCode=*QuickDel.VslCode AND " & _
				"mShipper.Shipper=*QuickDel.Shipper AND BL.BLNo=*QuickDel.BLNo AND " & _
				"QuickDel.Forwarder='" & sForwarder & "'"
	rsd.Open sql, conn, 0, 1, 1

	Dim ShipperAbrev(),ShipLineAbrev(),VslFull(),BLNo(),CntrNo(),RejectFlg(),RecSchTime()
	Dim Shipper(),ShipLine(),VslCode(),OpeCode()
	QdelNo=0
	Do While Not rsd.EOF
		ReDim Preserve ShipperAbrev(QdelNo)
		ReDim Preserve ShipLineAbrev(QdelNo)
		ReDim Preserve VslFull(QdelNo)
		ReDim Preserve BLNo(QdelNo)
		ReDim Preserve CntrNo(QdelNo)
		ReDim Preserve RejectFlg(QdelNo)
		ReDim Preserve RecSchTime(QdelNo)
		ReDim Preserve Shipper(QdelNo)
		ReDim Preserve ShipLine(QdelNo)
		ReDim Preserve VslCode(QdelNo)
		ReDim Preserve OpeCode(QdelNo)
		ShipperAbrev(QdelNo) = Trim(rsd(0))
		ShipLineAbrev(QdelNo) = Trim(rsd(1))
		VslFull(QdelNo) = Trim(rsd(2))
		BLNo(QdelNo) = Trim(rsd(3))
		CntrNo(QdelNo) = Trim(rsd(4))
		RejectFlg(QdelNo) = Trim(rsd(5))
		RecSchTime(QdelNo) = DispDateTime(rsd(6),0)
		Shipper(QdelNo) = Trim(rsd(7))
		ShipLine(QdelNo) = Trim(rsd(8))
		VslCode(QdelNo) = Trim(rsd(9))
		OpeCode(QdelNo) = Trim(rsd(10))
		QdelNo=QdelNo+1
	  rsd.MoveNext
	Loop
	rsd.Close

	Dim LineNo,OpeAbrev,OpeTelNo,strOut
	LineNo=0
	' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
	strFileName="./temp/" & strFileName
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

	For i=0 to QdelNo-1
		'' BL�����݂��Ȃ���΁A
		If BLNo(i) = "" Then
			sql = "SELECT BL.OpeCode, ImportCont.VslCode, ImportCont.VoyCtrl FROM BL,ImportCont " & _
						"WHERE BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl AND " & _
						"ImportCont.ContNo='" & CntrNo(i) & "' ORDER BY ImportCont.UpdtTime DESC"
			rsd.Open sql, conn, 0, 1, 1
			Do While Not rsd.EOF
				OpeCode(i) = Trim(rsd(0))
				Exit Do
				rsd.MoveNext
			Loop
			rsd.Close
		End If

		'' DB�̓ǂݍ���
		sql = "SELECT NameAbrev,TelNo FROM mOperator WHERE OpeCode='" & OpeCode(i) & "'"
		rsd.Open sql, conn, 0, 1, 1
		OpeAbrev=""
		OpeTelNo=""
		Do While Not rsd.EOF
			OpeAbrev = Trim(rsd(0))
			OpeTelNo = Trim(rsd(1))
			Exit Do
			rsd.MoveNext
		Loop
		rsd.Close

		strOut = ShipperAbrev(i) & ","
		strOut = strOut & ShipLineAbrev(i) & ","
		strOut = strOut & VslFull(i) & ","
		strOut = strOut & BLNo(i) & ","
		strOut = strOut & CntrNo(i) & ","
		strOut = strOut & OpeAbrev & ","
		strOut = strOut & OpeTelNo & ","
		If RejectFlg(i) = "0" then
			strOut = strOut & "��" & ","
		ElseIf RejectFlg(i) = "1" then
			strOut = strOut & "�~" & ","
		Else
			strOut = strOut & "" & ","
		End If
		strOut = strOut & RecSchTime(i) & ","
		strOut = strOut & Shipper(i) & ","
		strOut = strOut & ShipLine(i) & ","
		strOut = strOut & VslCode(i)
		ti.WriteLine strOut
		LineNo=LineNo+1
	Next
	ti.Close

	If LineNo>0 Then
		Response.Redirect "sokuji-kaika-list.asp"
		Response.End
'	Else
'		Response.Redirect "sokuji-kaika-new.asp?kind=1"
'		Response.End
	End If

%>

<html>
<head>
<title>�������o�\���ݏ��ꗗ�i�C�݁j</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function formSend(formname){
		window.document.forms[formname].submit();
	}

<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------�������烍�O�C�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/sokuji1t.gif" width="506" height="73"></td>
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

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
<%
	strNowTime = Year(Now) & "�N" & _
		Right("0" & Month(Now), 2) & "��" & _
		Right("0" & Day(Now), 2) & "��" & _
		Right("0" & Hour(Now), 2) & "��" & _
		Right("0" & Minute(Now), 2) & "�����݂̏��"

%>
					&nbsp;&nbsp;<%=strNowTime%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>�i�C�ݗp�j�������o�\���ݏ��ꗗ</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>

            <br>

			<table border=0 cellpadding=0>
			  <tr>
				<td align=center colspan=2>

					<table border=0 cellpadding=0 cellspacing=2>
					<tr><td>
					�׎�R�[�h��I������ƍX�V���\�ł��B 
					&nbsp;�V�K�̏ꍇ�́A'�V�K����' ���N���b�N���ĉ������B
					</td></tr>
					</table>
					<BR>
				</td>
			  </tr>
			  <tr>
				<td width=30><BR></td>
				<td nowrap>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>�׎�</td>
                <td nowrap>�D��</td>
                <td nowrap>�D��</td>
                <td nowrap>BL�^�R���e�iNo.</td>
                <td nowrap>�Ή��`�^</td>
                <td nowrap>�Ή��`�^<BR>TEL</td>
                <td nowrap>�Ή�<BR>��</td>
                <td nowrap>�����m�F�\�莞��</td>
              </tr>


			  <tr>
				<td nowrap align=center valign=middle>
					<a href="sokuji-kaika-new.asp?kind=1">�V�K����</a>
				</td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
			  </tr>
			</table>

			<form method=get action="sokuji-kaika-updtchk.asp">
				<input type=submit value="�\���f�[�^�̍X�V">
			</form>

				</td>
			  </tr>
			</table>

		  </td>
		</tr>
	  </table>

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
<!-------------���O�C����ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
%>
