<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>���o���\��\�����</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
function ClickSend() {

	/* ���o���^�C�v���擾 */
	Type1 = Type2 = Type3 = Type4 = 0
	for (i = 0; i < 3; i++) {
		if (document.SEND.RDType1[i].checked) {
			Type1 = i + 1
		}
	}
	for (i = 0; i < 3; i++) {
		if (document.SEND.RDType2[i].checked) {
			Type2 = i + 1
		}
	}
	for (i = 0; i < 3; i++) {
		if (document.SEND.RDType3[i].checked) {
			Type3 = i + 1
		}
	}
	for (i = 0; i < 3; i++) {
		if (document.SEND.RDType4[i].checked) {
			Type4 = i + 1
		}
	}

	if (Type1 == 0 &&
		Type2 == 0 &&
		Type3 == 0 &&
		Type4 == 0) {
		window.alert("��ނ���͂��Ă��������B");
		return false;
	}

	if (ChkSend("�\��P", Type1,
				document.SEND.ContNoRec1.value,
				document.SEND.BKNo1.value,
				document.SEND.ContSizeRec1.value,
				document.SEND.checkA1.checked,
				document.SEND.checkB1.checked,
				document.SEND.checkC1.checked,
				document.SEND.ReceiveFrom1.value,
				document.SEND.ContNoDel1.value,
				document.SEND.ChID1.value,
				document.SEND.BLNo1.value,
				document.SEND.ContSizeDel1.value,
				document.SEND.DeliverTo1.value,
				document.SEND.NinID1.value) &&
		ChkSend("�\��Q", Type2,
				document.SEND.ContNoRec2.value,
				document.SEND.BKNo2.value,
				document.SEND.ContSizeRec2.value,
				document.SEND.checkA2.checked,
				document.SEND.checkB2.checked,
				document.SEND.checkC2.checked,
				document.SEND.ReceiveFrom2.value,
				document.SEND.ContNoDel2.value,
				document.SEND.ChID2.value,
				document.SEND.BLNo2.value,
				document.SEND.ContSizeDel2.value,
				document.SEND.DeliverTo2.value,
				document.SEND.NinID2.value) &&
		ChkSend("�\��R", Type3,
				document.SEND.ContNoRec3.value,
				document.SEND.BKNo3.value,
				document.SEND.ContSizeRec3.value,
				document.SEND.checkA3.checked,
				document.SEND.checkB3.checked,
				document.SEND.checkC3.checked,
				document.SEND.ReceiveFrom3.value,
				document.SEND.ContNoDel3.value,
				document.SEND.ChID3.value,
				document.SEND.BLNo3.value,
				document.SEND.ContSizeDel3.value,
				document.SEND.DeliverTo3.value,
				document.SEND.NinID3.value) &&
		ChkSend("�\��S", Type4,
				document.SEND.ContNoRec4.value,
				document.SEND.BKNo4.value,
				document.SEND.ContSizeRec4.value,
				document.SEND.checkA4.checked,
				document.SEND.checkB4.checked,
				document.SEND.checkC4.checked,
				document.SEND.ReceiveFrom4.value,
				document.SEND.ContNoDel4.value,
				document.SEND.ChID4.value,
				document.SEND.BLNo4.value,
				document.SEND.ContSizeDel4.value,
				document.SEND.DeliverTo4.value,
				document.SEND.NinID4.value)) {
		return true;
	}
	return false;
}

function ChkSend(Name, RDType, ContNoRec, BKNo, ContSizeRec,
					ChkA, ChkB, ChkC, ReceiveFrom,
					ContNoDel, ChID, BLNo, ContSizeDel, DeliverTo, NinID) {
	if (RDType == 0) {					/*�I���Ȃ�*/
		if (ContNoRec != "" || BKNo != "" || ContSizeRec != "BL" || ReceiveFrom != "" ||
			ChkA || ChkB  || ChkC  ||
			ContNoDel != "" || ChID != "" || ContSizeDel != "BL" ||
			DeliverTo != "" || NinID != "") {
				window.alert(Name + "�̎�ނ�I�����Ă��������B" + DeliverTo);
				return false;
		}
	}
	if (RDType == 1 || RDType == 2) {	/* �����̏ꍇ */
		if (ContNoRec == "") {
			window.alert(Name + "�̔����R���e�i�ԍ�����͂��Ă��������B");
			return false;
		}
		if (BKNo == "") {
			window.alert(Name + "�̔����u�b�L���O�ԍ�����͂��Ă��������B");
			return false;
		}
		if (ContSizeRec == "BL") {
			window.alert(Name + "�̔����R���e�i�T�C�Y����͂��Ă��������B");
			return false;
		}
		if (!ChkChara(ReceiveFrom)) {
			window.alert(Name + "�̃R���e�i�������͉p���œ��͂��ĉ������B");
			return false;
		}
	}
	if (RDType == 1 || RDType == 3) {	/* ���o�̏ꍇ */
		if (ContNoDel == "" && BLNo == "") {
			window.alert(Name + "�̔��o�R���e�i�ԍ����a�k�ԍ��̂ǂ��炩����͂��Ă��������B");
			return false;
		}
		if (ContNoDel != "" && BLNo != "") {
			window.alert(Name + "�̔��o�R���e�i�ԍ����a�k�ԍ��̂ǂ��炩����͂��Ă��������B");
			return false;
		}
		if (BLNo != "") {	/* BL�w��̏ꍇ */
			if (ContSizeDel == "BL") {
				window.alert(Name + "�̔��o�R���e�i�T�C�Y����͂��Ă��������B");
				return false;
			}
		}
		if (!ChkChara(DeliverTo)) {
			window.alert(Name + "�̃R���e�i���o��͉p���œ��͂��ĉ������B");
			return false;
		}
	}
	if (RDType == 1) {				/* ���o���̏ꍇ */
		if (ChkA) {
			window.alert(Name + "�́w���o���ڂ��Ȃ��x�̓f���A���̏ꍇ�ɂ͖����ł��B");
			return false;
		}
		if (ChkB) {
			window.alert(Name + "�́w�[�ς̂ݍڂ���x�̓f���A���̏ꍇ�ɂ͖����ł��B");
			return false;
		}
		if (ChID != "") {
			window.alert(Name + "�̃V���[�VID�w��̓f���A���̏ꍇ�ɂ͖����ł��B");
			return false;
		}
		if (!ChkC && BLNo != "" &&
			((ContSizeRec != "20" && ContSizeDel == "20") ||
		     (ContSizeRec == "20" && ContSizeDel != "20"))) {
			window.alert(Name + "�̔����Ɣ��o�ŃR���e�i�T�C�Y���قȂ�܂��B");
			return false;
		}
	}
	if (ChkA && ChkB) {
		window.alert(Name + "�́w���o���ڂ��Ȃ��x�Ɓw�[�ς̂ݍڂ���x���������Ă��܂��B");
		return false;
	}
	return true;
}
function ChkChara(str) {
	/* ���p�p�������̂݋��� */
	sWk = str.toUpperCase()	/* �啶���ϊ� */
	for (i = 0; i < sWk.length; i++) {
		if (!((sWk.charAt(i) >= "A" && sWk.charAt(i) <= "Z") ||
 		      (sWk.charAt(i) >= "0" && sWk.charAt(i) <= "9"))) {
			return false;
		}
	}
	return true;
}
function ClickSend1(go) {
	/*�N���A���� �\��P*/
	for (i = 0; i < 3; i++) {
		document.SEND.RDType1[i].checked = false
	}
	document.SEND.ContNoRec1.value = ""
	document.SEND.BKNo1.value = ""
	document.SEND.ContSizeRec1.value = "BL"
	document.SEND.checkA1.checked = false
	document.SEND.checkB1.checked = false
	document.SEND.checkC1.checked = false
	document.SEND.ReceiveFrom1.value = ""
	document.SEND.ContNoDel1.value = ""
	document.SEND.ChID1.value = ""
	document.SEND.BLNo1.value = ""
	document.SEND.ContSizeDel1.value = "BL"
	document.SEND.DeliverTo1.value = ""
	document.SEND.NinID1.value = ""
}
function ClickSend2(go) {
	/*�N���A���� �\��Q*/
	for (i = 0; i < 3; i++) {
		document.SEND.RDType2[i].checked = false
	}
	document.SEND.ContNoRec2.value = ""
	document.SEND.BKNo2.value = ""
	document.SEND.ContSizeRec2.value = "BL"
	document.SEND.checkA2.checked = false
	document.SEND.checkB2.checked = false
	document.SEND.checkC2.checked = false
	document.SEND.ReceiveFrom2.value = ""
	document.SEND.ContNoDel2.value = ""
	document.SEND.ChID2.value = ""
	document.SEND.BLNo2.value = ""
	document.SEND.ContSizeDel2.value = "BL"
	document.SEND.DeliverTo2.value = ""
	document.SEND.NinID2.value = ""
}
function ClickSend3(go) {
	/*�N���A���� �\��R*/
	for (i = 0; i < 3; i++) {
		document.SEND.RDType3[i].checked = false
	}
	document.SEND.ContNoRec3.value = ""
	document.SEND.BKNo3.value = ""
	document.SEND.ContSizeRec3.value = "BL"
	document.SEND.checkA3.checked = false
	document.SEND.checkB3.checked = false
	document.SEND.checkC3.checked = false
	document.SEND.ReceiveFrom3.value = ""
	document.SEND.ContNoDel3.value = ""
	document.SEND.ChID3.value = ""
	document.SEND.BLNo3.value = ""
	document.SEND.ContSizeDel3.value = "BL"
	document.SEND.DeliverTo3.value = ""
	document.SEND.NinID3.value = ""
}
function ClickSend4(go) {
	/*�N���A���� �\��S*/
	for (i = 0; i < 3; i++) {
		document.SEND.RDType4[i].checked = false
	}
	document.SEND.ContNoRec4.value = ""
	document.SEND.BKNo4.value = ""
	document.SEND.ContSizeRec4.value = "BL"
	document.SEND.checkA4.checked = false
	document.SEND.checkB4.checked = false
	document.SEND.checkC4.checked = false
	document.SEND.ReceiveFrom4.value = ""
	document.SEND.ContNoDel4.value = ""
	document.SEND.ChID4.value = ""
	document.SEND.BLNo4.value = ""
	document.SEND.ContSizeDel4.value = "BL"
	document.SEND.DeliverTo4.value = ""
	document.SEND.NinID4.value = ""
}

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sHH, sHHName, sTerm_Name, sTerm_CD

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD, 9, 2)
	sYMD = Left(sYMD, 8)
	sHHName = TRIM(Request.QueryString("Name"))
	sTerm_Name = Trim(Request.QueryString("Term_Name"))		'VP�Ή�
	sTerm_CD = Trim(Request.QueryString("Terminal"))		'VP�Ή�
%>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title22.gif" width="236" height="34"><p>
</center>
<center>
<table border="1">
	<tr ALIGN=middle>
		<td width="200" bgcolor ="#e8ffe8">��Ǝ���</td>
		<td width="360" bgcolor ="#ffffff"><%=ChgYMDStr2(sYMD)%>�@<%=sHHName%></td>
	</tr>
	<tr ALIGN=middle>
		<td width="200" bgcolor ="#e8ffe8">���o����</td>
		<td width="360" bgcolor ="#ffffff"><%=sTerm_Name%></td>
	</tr>
</table>
<br>
<center><font color="#ff0000"><small>
�i���Ӂj�R���e�i�������i���o��j�͔��p���[�}���œ��͂��Ă�������
</small></font>
</center>
<font face="�l�r �S�V�b�N">
<!--	<form  METHOD="post" NAME="SEND" ACTION="SYWB012.asp?TDATE=<%=sYMD%>&HH=<%=sHH%>&HHNAME=<%=sHHName%>" onSubmit="return ClickSend()"> -->
<form  METHOD="post" NAME="SEND" ACTION="SYWB012.asp?TDATE=<%=sYMD%>&HH=<%=sHH%>&HHNAME=<%=sHHName%>
													&Term_Name=<%=sTerm_Name%>&Terminal=<%=sTerm_CD%>"
			onSubmit="return ClickSend()">
<center>
<%
	Dim idx, sRDType
	Dim sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC
	Dim sContNoDel, sChID, sBLNo, sContSizeDel, sDeliverTo, sReceiveFrom
'2003/08/27 �F��ID�̒ǉ�(ICCT�Ή�)
	Dim sWk, sNinID

	for idx = 1 to 4
'���L<br>���Ƃ���9/13
%>
	<table border="0" width="700" bgcolor ="#ffffff">
		<TR><th align=left><font color="#00008B">���\��<%=idx%>��</font></th></TR>
	</table>

	<table border="1" width="700" bgcolor ="#ffffff" cellpadding="3">
		<tr><th bgcolor ="#40E0D0">���</th>
			<td COLSPAN=2 bgcolor ="#ffffcc">
<%			sRDType = TRIM(Request.QueryString("sRDType" & CStr(idx)))
			If sRDType	= "" or sRDType	= null	then %>
				<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DUAL">DUAL
				<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="REC">����
				<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DEL">���o�@</td>
<%			Else
				Select case  sRDType
					case	"DUAL"	%>
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DUAL" Checked>DUAL
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="REC">����
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DEL">���o�@</td>
<%					case	"REC"	%>
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DUAL">DUAL
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="REC" Checked>����
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DEL">���o�@</td>
<%					case	"DEL"	%>
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DUAL">DUAL
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="REC">����
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="DEL" Checked>���o�@</td>
<%				End Select
			End If							%>
		</tr>
		<tr><th bgcolor ="#40E0D0" ROWSPAN=2>������</th>
<%			If	sRDType = "" OR sRDType = "DEL" Then 							%>
				<td bgcolor=#cccc99>
					�R���e�i�ԍ��@(�K�{)<INPUT TYPE="text" NAME="ContNoRec<%=CStr(idx)%>" SIZE="18" MAXLENGTH="12"><br>
				    �u�b�L���O�ԍ�(�K�{)<INPUT TYPE="text" NAME="BKNo<%=CStr(idx)%>" SIZE="28" MAXLENGTH="20"><br>
					�R���e�i�T�C�Y(�K�{)<SELECT NAME="ContSizeRec<%=CStr(idx)%>" size=0>
									<OPTION VALUE="BL" selected>
									<OPTION VALUE="20" >20
									<OPTION VALUE="40" >40</OPTION>
								</SELECT></td>
				<td bgcolor ="#ffffcc">
					<INPUT TYPE=checkbox NAME="checkA<%=CStr(idx)%>"> ���o���ڂ��Ȃ�(�I��)<br>
					<INPUT TYPE=checkbox NAME="checkB<%=CStr(idx)%>"> �[�ς̂ݍڂ���(�I��)<br>
					<INPUT TYPE=checkbox NAME="checkC<%=CStr(idx)%>"> 20/40���p�V���[�V(�I��)
				</td>
				</tr>
				<tr>
				<td colspan=3 bgcolor ="#ffffcc">
				(����)���m�̏ꍇ�E�E�E�R���e�i������
					<INPUT NAME="ReceiveFrom<%=CStr(idx)%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
				</td>
				</tr>

<%			Else
				sContNoRec = TRIM(Request.QueryString("sContNoRec" & CStr(idx)))
				sBKNo = UCASE(TRIM(Request.QueryString("sBKNo" & CStr(idx))))
				sContSizeRec = UCASE(TRIM(Request.QueryString("sContSizeRec" & CStr(idx))))
				bChkA = Request.QueryString("bChkA" & CStr(idx))
				bChkB = Request.QueryString("bChkB" & CStr(idx))
				bChkC = Request.QueryString("bChkC" & CStr(idx))
				sReceiveFrom = Leftb(TRIM(Request.QueryString("sReceiveFrom" & CStr(idx))),30)
%>
				<td bgcolor=#cccc99>
					�R���e�i�ԍ��@(�K�{)<INPUT TYPE="text" NAME="ContNoRec<%=CStr(idx)%>" Value="<%=sContNoRec%>" SIZE="18" MAXLENGTH="12"><br>
				    �u�b�L���O�ԍ�(�K�{)<INPUT TYPE="text" NAME="BKNo<%=CStr(idx)%>" Value="<%=sBKNo%>" SIZE="28" MAXLENGTH="20"><br>
					�R���e�i�T�C�Y(�K�{)<SELECT NAME="ContSizeRec<%=CStr(idx)%>" size=0>
<%					Select Case	sContSizeRec
						Case	"20"	%>
								<OPTION VALUE="BL" >
								<OPTION VALUE="20" selected>20
								<OPTION VALUE="40" >40</OPTION>
							</SELECT></td>
<%						Case	"40"	%>
								<OPTION VALUE="BL" >
								<OPTION VALUE="20" >20
								<OPTION VALUE="40" selected>40</OPTION>
							</SELECT></td>
<%					End Select			%>
				<td bgcolor ="#ffffcc">
<%					If bChkA = "True" Then	%>
						<INPUT TYPE=checkbox NAME="checkA<%=CStr(idx)%>" Checked> ���o���ڂ��Ȃ�(�I��)<br>
<%					Else					%>
						<INPUT TYPE=checkbox NAME="checkA<%=CStr(idx)%>"> ���o���ڂ��Ȃ�(�I��)<br>
<%					End If					%>

<%					If bChkB = "True" Then	%>
						<INPUT TYPE=checkbox NAME="checkB<%=CStr(idx)%>" Checked> �[�ς̂ݍڂ���(�I��)<br>
<%					Else					%>
						<INPUT TYPE=checkbox NAME="checkB<%=CStr(idx)%>"> �[�ς̂ݍڂ���(�I��)<br>
<%					End If					%>

<%					If bChkC = "True" Then	%>
						<INPUT TYPE=checkbox NAME="checkC<%=CStr(idx)%>" Checked> 20/40���p�V���[�V(�I��)
<%					Else					%>
						<INPUT TYPE=checkbox NAME="checkC<%=CStr(idx)%>"> 20/40���p�V���[�V(�I��)
<%					End If					%>
				</td>
				</tr>
				<tr>
				<td colspan=3 bgcolor ="#ffffcc">
<%				If sReceiveFrom <> "" Then	%>
				(����)���m�̏ꍇ�E�E�E�R���e�i�������@<INPUT NAME="ReceiveFrom<%=CStr(idx)%>" Value="<%=sReceiveFrom%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
<%				Else%>
				(����)���m�̏ꍇ�E�E�E�R���e�i�������@<INPUT NAME="ReceiveFrom<%=CStr(idx)%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
<%				End If	%>
				</td>
				</tr>
<%			End If	%>

		<tr><th bgcolor ="#40E0D0" ROWSPAN=3>���o��</th>
<%		If sRDType	= "DEL" or sRDType	= "DUAL"	then
			sContNoDel = TRIM(Request.QueryString("sContNoDel" & CStr(idx)))
			sChID = TRIM(Request.QueryString("sChID" & CStr(idx)))
            sNinID = TRIM(Request.QueryString("sNinID" & CStr(idx)))
			sBLNo = TRIM(Request.QueryString("sBLNo" & CStr(idx)))
			sContSizeDel = TRIM(Request.QueryString("sContSizeDel" & CStr(idx)))
			sDeliverTo = Leftb(TRIM(Request.QueryString("sDeliverTo" & CStr(idx))),30)
%>
			<td bgcolor ="#cccc99" COLSPAN=2>�R���e�i�ԍ�
				<INPUT TYPE="text" NAME="ContNoDel<%=CStr(idx)%>" Value="<%=sContNoDel%>" SIZE="18" MAXLENGTH="12">
			</td>
			</tr>
			<tr><td bgcolor ="#cccc99" COLSPAN=2>�܂��́A�a�k�ԍ�
				<INPUT NAME="BLNo<%=CStr(idx)%>" Value="<%=sBLNo%>" SIZE="28" MAXLENGTH="20">
<%
			If sBLNo <> "" Then
				Select Case sContSizeDel
					Case "20"			%>
				�T�C�Y(�K�{)<SELECT NAME="ContSizeDel<%=CStr(idx)%>" size=0>
								<OPTION VALUE="BL" >
								<OPTION VALUE="20" selected>20
								<OPTION VALUE="40" >40</OPTION>
							</SELECT>
<%					Case "40"		%>
				�T�C�Y(�K�{)<SELECT NAME="ContSizeDel<%=CStr(idx)%>" size=0>
								<OPTION VALUE="BL" >
								<OPTION VALUE="20" >20
								<OPTION VALUE="40" selected>40</OPTION>
							</SELECT>
<%				End Select
			Else						%>
				�T�C�Y(�K�{)<SELECT NAME="ContSizeDel<%=CStr(idx)%>" size=0>
								<OPTION VALUE="BL" selected>
								<OPTION VALUE="20" >20
								<OPTION VALUE="40" >40</OPTION>
							</SELECT>
<%			End IF						%>
			</td>
			</tr>
			<tr>
			<td COLSPAN=2 bgcolor ="#ffffcc">
				(����)���m�̏ꍇ�E�E�E�R���e�i���o��@
				<INPUT NAME="DeliverTo<%=CStr(idx)%>" Value="<%=sDeliverTo%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
				(����)�K�v�ɉ����āE�E�E�V���[�V�h�c �@<INPUT NAME="ChID<%=CStr(idx)%>" Value="<%=sChID%>" SIZE="9" MAXLENGTH="5"><br>
				(����)�K�v�ɉ����āE�E�E�F�؂h�c �@�@�@<INPUT NAME="NinID<%=CStr(idx)%>" Value="<%=sNinID%>" SIZE="18" MAXLENGTH="10"></td>
			</tr>
<%
		Else
		%>
			<td bgcolor ="#cccc99" COLSPAN=2 >�R���e�i�ԍ�
			<INPUT TYPE="text" NAME="ContNoDel<%=CStr(idx)%>" SIZE="18" MAXLENGTH="12">
			</td>
			</tr>
			<tr><td bgcolor ="#cccc99" COLSPAN=2>�܂��́A�a�k�ԍ�
				<INPUT NAME="BLNo<%=CStr(idx)%>" SIZE="28" MAXLENGTH="20">
				�T�C�Y(�K�{)<SELECT NAME="ContSizeDel<%=CStr(idx)%>" size=0>
								<OPTION VALUE="BL" selected>
								<OPTION VALUE="20" >20
								<OPTION VALUE="40" >40</OPTION>
							</SELECT>
			</td>
			</tr>
			<tr>
			<td COLSPAN=2 bgcolor ="#ffffcc">(����)���m�̏ꍇ�E�E�E�R���e�i���o��@
			<INPUT NAME="DeliverTo<%=CStr(idx)%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
				(����)�K�v�ɉ����āE�E�E�V���[�V�h�c �@<INPUT NAME="ChID<%=CStr(idx)%>" SIZE="9" MAXLENGTH="5"><br>
				(����)�K�v�ɉ����āE�E�E�F�؂h�c �@�@�@<INPUT NAME="NinID<%=CStr(idx)%>" SIZE="18" MAXLENGTH="10"></td>
			</tr>
<%
		End If	%>
	</table>

	<table border=0 width="700" bgcolor ="#ffffff">
		<tr><td align=center><font color="#ff0000"><small>
		�i���Ӂj�R���e�i������(���o��)�̓_�C������܂łɓ��͂��Ȃ��ꍇ�\��L�����Z���ƂȂ�܂�
		</small></font></td></tr>
	</table>

	<table border=0 width="700" bgcolor ="#ffffff"><tr align=right><td>
		<input type="submit" value="�@�S�̑��M�@" id=submit4 name=submit4>
		<input type="button" value="�\��<%=CStr(idx)%>�ر" id=submit4 name=submit4 onclick="return ClickSend<%=CStr(idx)%>(this)">
		</td></tr>
	</table>

<%	next	%>
</center>

<br>

<center>
<table border=0>
	</form>

    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="�@�@���~�@�@" id=submit4 name=submit4></td>
	</form>
</table>
</center>

</body>
</html>
