<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB033.inc"-->
<html>

<head>
<title>�V���[�V�\��ύX�o�^���</title>
</head>
<body>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title27.gif" width="236" height="34"><p>
</center>

		<font face="�l�r �S�V�b�N">
<CENTER>
<%
	Dim conn, rsd, sql									'�c�a�ڑ�
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator	'���[�U���
	Dim sYMD, iOpeNo									'�w����t�A��Ɣԍ�
	Dim sM_ChassisId, sChkChassisID						'���̃V���[�V�h�c�A�ύX�i�����j�V���[�V�h�c
	Dim sHK_flg											'�����t���O�iK�F�����j
	Dim sErr_msg										'�G���[���b�Z�[�W
 	Dim iChg_OpeNo										'���������Ɣԍ�
	Dim sSize20Flag_O, sMixSizeFlag_O, sGroupID_O		'���V���[�V����
	Dim sSize20Flag_N, sMixSizeFlag_N, sGroupID_N		'�V�V���[�V����
	Dim iDualOpe, sHH									'�f���A����Ɣԍ��A���ԑ�
	Dim iOpeOrder										'��Ə���
	Dim sBlanks,sBlanke									'�󔒍s(�G���[���̉�ʒ����p)


	'�G���[���̉�ʒ����p
	sBlanks = "<B><U><font color=#ff0000><br><br><br><br>"
	sBlanke = "<br><br><br><br><br><br></font></U></B>"

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))

	'��Ɣԍ��擾
	iOpeNo = TRIM(Request.QueryString("OPENO"))

	'���̃V���[�V�h�c�擾
	sM_ChassisId = TRIM(Request.QueryString("M_ChassisId"))

	'�ύX�i�����j�V���[�V�h�c�̎擾
	If Len(TRIM(Request.QueryString("sCassis"))) = 5 Then
		sChkChassisID = Left(TRIM(Request.QueryString("sCassis")),4)
	Else
		sChkChassisID = Left(TRIM(Request.QueryString("sCassis")),5)
	End If

	'�����t���O�擾
	sHK_flg = Right(TRIM(Request.QueryString("sCassis")),1)

	'���̃V���[�V�h�c�̑����擾
	If sM_ChassisId <> "" Then
		sql = "SELECT Size20Flag, MixSizeFlag, GroupID FROM sChassis" & _
				" WHERE ChassisId = '" & sM_ChassisId & "'"
		rsd.Open sql, conn, 0, 1, 1
		If	rsd.EOF Then		'���R�[�h���Ȃ��ꍇ
			rsd.close
			sErr_msg = sBlanks & "���͂��ꂽ�V���[�V�͑��݂��܂���B" & sBlanke
			Response.Write sErr_msg
%><center><input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()"><center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If
		sSize20Flag_O  = Trim(rsd("Size20Flag"))		'�Q�O�t�B�[�g�t���O
		sMixSizeFlag_O = Trim(rsd("MixSizeFlag"))		'�Q�O�^�S�O���p�V���[�V
		sGroupID_O	   = Trim(rsd("GroupID"))			'�O���[�v�h�c
		rsd.close
	End If

	'�ύX�V���[�V�h�c�̑����擾
	sql = "SELECT Size20Flag, MixSizeFlag, GroupID FROM sChassis" & _
			" WHERE ChassisId = '" & sChkChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1
	If	rsd.EOF Then		'���R�[�h���Ȃ��ꍇ
		rsd.close
		sErr_msg = sBlanks & "���͂��ꂽ�V���[�V�͑��݂��܂���B" & sBlanke
		Response.Write sErr_msg
		%><center>
		<input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()" id=button1 name=button1>
		<center><%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End If
	sSize20Flag_N  = Trim(rsd("Size20Flag"))		'�Q�O�t�B�[�g�t���O
	sMixSizeFlag_N = Trim(rsd("MixSizeFlag"))		'�Q�O�^�S�O���p�V���[�V
	sGroupID_N	   = Trim(rsd("GroupID"))			'�O���[�v�h�c
	rsd.close

	'�ύX�i�����j�O�㑮���`�F�b�N����
	If sM_ChassisId <> "" Then	'���V���[�V������ꍇ
		'�V���[�V�T�C�Y�̕s�K�����`�F�b�N
		If sSize20Flag_O <> sSize20Flag_N Or _
		   sMixSizeFlag_O <> sMixSizeFlag_N Then
			sErr_msg = sBlanks & "���͂��ꂽ�V���[�V�͏����ɂ����܂���B�P" & sBlanke
			Response.Write sErr_msg
			%><center>
			<input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()" id=button2 name=button2>
			<center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End if
		'�O���[�v�̃`�F�b�N���s��
		If sGroupID_O <> sGroupID_N then
			sErr_msg = sBlanks & sM_ChassisId & "��" & _
						sChkChassisID & "�@�̃V���[�V�͕ύX�i�����j�ł��܂���B" & sBlanke
			Response.Write sErr_msg
			%><center>
			<input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()" id=button3 name=button3>
			<center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End if
	End If

	'������̍�Ɣԍ����擾����
	if sHK_flg = "K" then	'������
		'�����Ɠ��ɗ\�񒆂̔��o�\�����擾
		sql = "SELECT distinct OpeNo FROM sAppliInfo"
		sql = sql & " WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'"
		sql = sql & "  AND Status   = '02'"
		sql = sql & "  AND RecDel   = 'D'"
		sql = sql & "  AND DelFlag  = ' '"
		sql = sql & "  AND WorkFlag = ' '"
		sql = sql & "  AND LockFlag = ' '"
		sql = sql & "  AND sAppliInfo.WorkDate = '" & cdate(ChgYMDStr(sYMD)) & "'"
		sql = sql & "  AND RTRIM(ChassisID) = '" & sChkChassisID & "'"
		rsd.Open sql, conn, 0, 1, 1
		If rsd.EOF Then	
			rsd.close
			sErr_msg = sBlanks & "��������̗\��̓V���[�V�����s�ł��B" & sBlanke
			Response.Write sErr_msg	
			%><center>
			<input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()" id=button5 name=button5>
			<center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If
		iChg_OpeNo = rsd("OpeNo")			'���������Ɣԍ�
		rsd.close
	End If

	'�ύX���\�����̎擾
	Call GetOApp(conn, rsd, iOpeNo, sYMD, sErr_msg)
	If sErr_msg <> "" Then	'�G���[�̂���ꍇ
		rsd.close
		Response.Write sBlanks & sErr_msg & sBlanke
		%><center><input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()" id=button6 name=button6>
		<center><%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End If

	'�V���[�V�T�C�Y�̕s�K�����`�F�b�N(�\�����̃T�C�Y���m�F����)
	If rsd("ContSize") = "20" Then 
'''		If sSize20Flag_N = "Y" Or rsd("MixSizeFlag") = "Y" Then	'2001/06/02
		If sSize20Flag_N = "Y" Or sMixSizeFlag_N = "Y" Then		'2001/06/02
			sErr_msg = ""
		Else
			sErr_msg = sBlanks & "���͂��ꂽ�V���[�V�͏����ɂ����܂���B" & sBlanke
		End If
	Else
		If sSize20Flag_N = "Y" Then
			sErr_msg = sBlanks & "���͂��ꂽ�V���[�V�͏����ɂ����܂���B" & sBlanke
		Else
			sErr_msg = ""
		End If
	End If

	If sErr_msg <> "" then
		Response.Write sErr_msg
		%><center>
		<input type="button" value="�@�m�F�@" onClick="JavaScript:history.back()" id=button4 name=button4>
		<center><%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End if

	iDualOpe  = rsd("DualOpeNo")		'�f���A����Ɣԍ�
	sHH       = Trim(rsd("Term"))		'���ԑ�
	iOpeOrder = rsd("OpeOrder")			'��Ə���
	rsd.close

	If iDualOpe > 0 Then	'�f���A���̏ꍇ�̓f���A����������
		'�V�K��Ə��ʂ̎擾�i�w����A�w�莞�ԑсj
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, "D")
	End If

	'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
	Call GetAppInfoOpeNoUpd(conn, rsd, iOpeNo)

	'�ύX���\�����̍X�V
	rsd("UpdtTime")  = now()			'�X�V��
	rsd("UpdtPgCd")  = "SYWB033"		'�X�V�v���O����
	rsd("ChassisID") = sChkChassisID	'�V���[�VID
	rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
	rsd("OpeOrder")  = iOpeOrder		'��Ə���
	rsd("SendFlag")  = "Y"				'���M�t���O
	rsd.update
	rsd.close

	If iDualOpe > 0 Then	'�f���A���̏ꍇ�̓f���A����������
		'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpe)
		If Not rsd.EOF Then		'�{�����R�[�h�͕K������
			'�ύX���\�����̍X�V
			rsd("UpdtTime")  = now()			'�X�V��
			rsd("UpdtPgCd")  = "SYWB033"		'�X�V�v���O����
			rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
			rsd("SendFlag")  = "Y"				'���M�t���O
			rsd.update
		End If
		rsd.close
	End If

	'�����̎�
	If sHK_flg = "K" then	'������
		'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
		Call GetAppInfoOpeNoUpd(conn, rsd, iChg_OpeNo)
		If Not rsd.EOF Then		'�{�����R�[�h�͕K������
			iDualOpe  = rsd("DualOpeNo")		'�f���A����Ɣԍ�
			sHH       = Trim(rsd("Term"))		'���ԑ�
			iOpeOrder = rsd("OpeOrder")			'��Ə���
			rsd.close
			If iDualOpe > 0 Then	'�f���A���̏ꍇ�̓f���A����������
				'�V�K��Ə��ʂ̎擾�i�w����A�w�莞�ԑсj
				iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, "D")
			End If

			'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
			Call GetAppInfoOpeNoUpd(conn, rsd, iChg_OpeNo)

			'�����\�����̍X�V
			rsd("UpdtTime")  = now()			'�X�V��
			rsd("UpdtPgCd")  = "SYWB033"		'�X�V�v���O����
			rsd("ChassisID") = sM_ChassisId		'�V���[�VID
			rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
			rsd("OpeOrder")  = iOpeOrder		'��Ə���
			rsd("SendFlag")  = "Y"				'���M�t���O
			rsd.update
			rsd.close
			If iDualOpe > 0 Then	'�f���A���̏ꍇ�̓f���A����������
				'�\�����̎擾�i�w���Ɣԍ��A�X�V���[�h�j
				Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpe)
				If Not rsd.EOF Then		'�{�����R�[�h�͕K������
					'�ύX���\�����̍X�V
					rsd("UpdtTime")  = now()			'�X�V��
					rsd("UpdtPgCd")  = "SYWB033"		'�X�V�v���O����
					rsd("DualOpeNo") = 0				'�f���A����Ɣԍ�
					rsd("SendFlag")  = "Y"				'���M�t���O
					rsd.update
				End If
				rsd.close
			End If
		Else
			rsd.close
		End if
	End if
%>
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
