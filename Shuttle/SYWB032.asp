<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB032.inc"-->
<html>

<head>
<title>�V���[�V�����o�^���</title>
</head>
<body>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title26.gif" width="236" height="34"><p>
</center>

		<font face="�l�r �S�V�b�N">
<%
	Dim conn, rsd, sql									'�c�a�ڑ�
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator	'���[�U���
	Dim sYMD, sHH										'�w����t
	Dim sChkChassisID									'�V���[�VID
	Dim s_chg_GrpID, s_chg_UsrID						'�ύX��O���[�v�A���[�U
	Dim sNotDelFlag										'����o�R���e�i���ڂ��Ȃ���w��iY�F�I���j
	Dim sNightFlag										'��[�ς݂̂ݍڂ��风w��iY�F�I���j
	Dim iOpeNo, iDualOpeNo								'��Ɣԍ��A�f���A����Ɣԍ�
	Dim iOpeOrder										'��Ə���
	Dim sWk

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'�w����t�擾
	sYMD = TRIM(Request.Form("YMD"))

	'�o�^�`�F�b�N
	If Request.Form("sy_zaiko")  <> "" Then
		sChkChassisID = Request.Form("sy_zaiko")	'�����
	ElseIf Request.Form("SELECT1")  <> "No0" Then
		sChkChassisID = Request.Form("SELECT1")		'�݌ɑI��
	Else
		sChkChassisID = Request.Form("SELECT2")		'��݌ɑI��
	End If

	sNotDelFlag = ""	'����o�R���e�i���ڂ��Ȃ���w��iY�F�I���j
	sNightFlag = ""		'��[�ς݂̂ݍڂ��风w��iY�F�I���j
	If Request.Form("check1") = "on" Then
		sNotDelFlag = "Y"	'����o�R���e�i���ڂ��Ȃ���w��iY�F�I���j
	End If
	If Request.Form("check2") = "on" Then
		sNightFlag = "Y"	'��[�ς݂̂ݍڂ��风w��iY�F�I���j
	End If

	'�O���[�v�ύX���̏���
	if Request.Form("check3") = "on" then		'�O���[�v�ύX�̏ꍇ
		'�w��V���[�V�̎g�p�\�肪���邩�`�F�b�N����
		If Not ChkAppCha(conn, rsd, sChkChassisID) Then
			%><center><%
			Response.Write sChkChassisID
			Response.Write "�@�̃V���[�V�͗\�񂳂�Ă���̂ő��̃O���[�v�ɂ͕ύX�ł��܂���B</p>"
			%><A HREF="JavaScript:history.back()">
				<BR>�V���[�V������ʂ֖߂�</A></CENTER> <%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If

		'�ύX��O���[�v�E���[�U�̎擾
		s_chg_GrpID = trim(Request.Form("SELECT3"))		'�O���[�v�R�[�h
		sql = "SELECT UserID,GroupID FROM sMUserGroup" & _
		  " WHERE RTRIM(GroupID) = '" & s_chg_GrpID & "'"
		rsd.Open sql, conn, 0, 1, 1
		if not rsd.eof then
			s_chg_UsrID = rsd("UserID")					'���[�U�R�[�h
		End If
		rsd.close	
	End If

	'�V���[�V����
	sql = "SELECT * FROM sChassis" & _
	  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
	  "   AND ChassisId = '" & sChkChassisID & "'"
	rsd.Open sql, conn, 0, 2, 1

	If rsd.EOF Then	
		rsd.close
		%><center><%
		Response.Write sChkChassisID
		Response.Write "�@�̃V���[�V�͑��݂��܂���B</p>"	
		%><A HREF="JavaScript:history.back()">
			<BR>�V���[�V������ʂ֖߂�</A></CENTER> <%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End If

	'�O���[�v�ύX�̏ꍇ�ɃV���[�V�̌�����`�F�b�N
	If Request.Form("check3") = "on" Then	'�O���[�v�ύX
		If rsd("ContFlag") = "Y" Then	'�R���e�i�t���O
			rsd.close
			%><center><%
			Response.Write sChkChassisID
			Response.Write "�@�̃V���[�V�ɂ̓R���e�i������܂��B</p>"
			%><A HREF="JavaScript:history.back()">
				<BR>�V���[�V������ʂ֖߂�</A></CENTER> <%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If

		If rsd("StackFlag") = "W" Then	'�V���g����ƒ�
			rsd.close
			%><center><%
			Response.Write sChkChassisID
			Response.Write "�@�̃V���[�V�̓V���g����ƒ��ł��B</p>"	
			%><A HREF="JavaScript:history.back()">
				<BR>�V���[�V������ʂ֖߂�</A></CENTER> <%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If
	End If

	rsd("UpdtTime") = now()				'�X�V��
	rsd("UpdtPgCd") = "SYWB0032"		'�X�V�v���O����

	rsd("NotDelFlag") = sNotDelFlag		'����o�R���e�i���ڂ��Ȃ���w��iY�F�I���j
	rsd("NightFlag")  = sNightFlag		'��[�ς݂̂ݍڂ��风w��iY�F�I���j
	If Request.Form("check3") = "on" Then	'��O���[�v�ύX��w��
		rsd("GroupID") = s_chg_GrpID	'�O���[�v�R�[�h
	End If
	rsd("SendFlag") = "Y"				'���M�t���O
	rsd.update
	rsd.close

	'�ΏۃV���[�V���g�p���̔������擾
	Call GetChangeApp(conn, rsd, sChkChassisID, iOpeNo, iDualOpeNo)
	If iOpeNo > 0 Then		'��Ƃ���
		'*** ������������ ***
		'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
		Call GetAppInfoOpeNoUpd(conn, rsd, iOpeNo)
		sHH = Trim(rsd("Term"))

		'�����\�����̍X�V
		rsd("UpdtTime")  = now()			'�X�V��
		rsd("UpdtPgCd")  = "SYWB032"		'�X�V�v���O����

		rsd("NotDelFlag") = sNotDelFlag		'����o�R���e�i���ڂ��Ȃ���w��iY�F�I���j
		rsd("NightFlag")  = sNightFlag		'��[�ς݂̂ݍڂ��风w��iY�F�I���j

		'����o�R���e�i���ڂ��Ȃ���w�肠�邢��'��[�ς݂̂ݍڂ��风w��̏ꍇ
		If sNotDelFlag = "Y" Or _
		   sNightFlag  = "Y" Then
			rsd("DualOpeNo") = 0			'�f���A����Ɣԍ�
		End If
		rsd("SendFlag")  = "Y"				'���M�t���O
		rsd.update
		rsd.close

		'*** ���o�������� ***
		If iDualOpeNo > 0 And _
		   (sNotDelFlag = "Y" Or _
		    sNightFlag  = "Y") Then
			'�V�K��Ə��ʂ̎擾�i�w����A�w�莞�ԑсj
			iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, "D")

			'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
			Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
			'�\�����̍X�V
			rsd("UpdtTime")  = now()			'�X�V��
			rsd("UpdtPgCd")  = "SYWB033"		'�X�V�v���O����
			rsd("ChassisID") = ""				'�V���[�VID
			rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
			rsd("OpeOrder")  = iOpeOrder		'��Ə���
			rsd("SendFlag")  = "Y"				'���M�t���O
			rsd.update
			rsd.close
		End If

		If iDualOpeNo > 0 Then	'�f���A���̏ꍇ
			sWk = "�c�t�`�k������܂��B���Ԙg�{�����z����\��������܂��B�_�C���m�莞�ɂ����ӂ�������"
		Else
			sWk = "�ݒ肵�܂����B��V���[�V�̉ߕs�����m�F���Ă��������B"
		End If
		%><CENTER><%=sWk%>
		  <A HREF=SYWB013.asp?TDATE=<%=sYMD%>>
				<BR>�ꗗ��ʂ֖߂�</A></CENTER>
		</body></html><%
		Response.end
	End If

%>
<CENTER>
<B>�X�V��</B>
</CENTER>
<FORM NAME="SEND">
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB013.asp?TDATE=" + document.SEND.YMD.value);
</SCRIPT>

</body>
</html>
 