<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>��o���\��ύX���</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
/* �ړ��̃N���b�N */
function ClickMov() {
	/* �`�F�b�N�Ȃ� */
	return true;
}
/* �����̃N���b�N */
function ClickChg() {

	if (document.CHG.CHGOPE.value == "") {
		window.alert("�����Ώۂ���͂��Ă��������B");
		return false;
	}
	if (document.CHG.CHGOPE.value.length != 5 && document.CHG.CHGOPE.value.length != 4) {
		window.alert("�����Ώۂ𐳂������͂��Ă��������B");
		return false;
	}length 
	return true;
}
/* ���o��̃N���b�N */
function ClickDel() {

		if (!ChkChara(document.UPD.DeliverTo.value)) {
			window.alert("�R���e�i���o��͔��p���[�}���œ��͂��ĉ������B");
			return false;
		}
}
function ClickRec() {

		if (!ChkChara(document.UPD.ReceiveFrom.value)) {
			window.alert("�R���e�i�������͔��p���[�}���œ��͂��ĉ������B");
			return false;
		}
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

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sHH, sHHName, sOpeNo, sAppTerminal
	Dim conn, rsd
	Dim sShtStart, sShtEnd, iSTime, iETime
	Dim iCnt, i
	Dim iTimeCnt, TimeSlot(40), TimeName(40)
	Dim sVPBookNo, sRecDel, sContSize
	Dim sDeliverTo, sReceiveFrom

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD, 9, 2)
	sYMD = Left(sYMD, 8)
	sHHName = TRIM(Request.QueryString("NAME"))

	'��Ɣԍ��擾
	sOpeNo = TRIM(Request.QueryString("OPENO"))

	'���o����擾
	sAppTerminal = TRIM(Request.QueryString("TNAME"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'�V���g���^�s���Ԏ擾
	sShtStart = GetEnv(conn, rsd, "ShtStart")
	sShtEnd   = GetEnv(conn, rsd, "ShtEnd")
	iSTime = CLng(Left(sShtStart, 2))
	iETime = CLng(Left(sShtEnd, 2))
	if Right(sShtEnd, 2) = "00" Then
		iETime = iETime - 1
	End If

	'�V���g���^�s���ԑьv�Z
	iCnt = 0

	'���ԑт̌v�Z
	''�ߑO����
	For i = iSTime To 11
		TimeSlot(iCnt) = Right("0" & CStr(i), 2)
		If i = iSTime Then
			TimeName(iCnt) = GetTimeSlot(i, CLng(Right(sShtStart, 2)), "S")
		Else
			TimeName(iCnt) = GetTimeSlot(i, "00", "S")
		End If
		iCnt = iCnt + 1
	Next
	''�ߑO�w��
	TimeSlot(iCnt) = "12"
	TimeName(iCnt) = "�ߑO"
	iCnt = iCnt + 1
	''�ߌ㎞��
	For i = 13 To iETime
		TimeSlot(iCnt) = Right("0" & CStr(i), 2)
		If i = iETime Then
			TimeName(iCnt) = GetTimeSlot(i + 1, CLng(Right(sShtEnd, 2)), "E")
		Else
			TimeName(iCnt) = GetTimeSlot(i + 1, "00", "E")
		End If
		iCnt = iCnt + 1
	Next
	''�ߌ�w��
	TimeSlot(iCnt) = "A"
	TimeName(iCnt) = "�ߌ�"
	iCnt = iCnt + 1
	''�[�ώw��
	TimeSlot(iCnt) = "B"
	TimeName(iCnt) = "�[��"
	iCnt = iCnt + 1

	iTimeCnt = iCnt		'���ԑѐ�

	'�c�a�擾
	Call GetAppInfoOpeNo(conn, rsd, int(sOpeNo))
	If Not rsd.EOF Then
		sVPBookNo = Trim(rsd("VPBookNo"))
		sRecDel = "��o��"
		sContSize = Trim(rsd("ContSize")) & "ft"  '�R���e�i�T�C�Y

'''		sReceiveFrom = Trim(rsd("ReceiveFrom"))			'������
		sDeliverTo = Trim(rsd("DeliverTo"))			'���o��
	end if
	rsd.Close

	'�c�a�ؒf
	conn.Close
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title35.gif" width="236" height="34"><p>
<table border="1">   
	<tr ALIGN=middle>
		<td width="120" bgcolor ="#e8ffe8">�Ώ�</td>
		<td width="380" ><%=ChgYMDStr2(sYMD)%>�@<%=sHHName%>�@<%=sOpeNo%><br>
			<%=sVPBookNo%>�@<%=sRecDel%>�@<%=sContSize%>�@<%=sAppTerminal%></td>
	</tr>
</table>
</center>
<br>

<font face="�l�r �S�V�b�N">

<center>

<form  METHOD="post" NAME="UPD" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=UPD&OPENO=<%=sOpeNo%>" onSubmit="return ClickDel()">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">���o��ύX</td></tr>
	<td></td>
	
		<td>
			<INPUT NAME="DeliverTo" Value="<%=sDeliverTo%>"	SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled">
			<input type="submit" value="�@���s�@" id=submit0 name=submit0>
		</td>
	</tr>
</table>
</form>

<form  METHOD="post" NAME="DEL" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=DEL&OPENO=<%=sOpeNo%>">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">�폜</td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<input type="submit" value="�@�폜�@" id=submit1 name=submit1>
		</td>
	</tr>
</table>
</form>
</center>

<center>
<form  METHOD="post" NAME="MOV" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=MOV&OPENO=<%=sOpeNo%>" onSubmit="return ClickMov()">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">�ړ�</td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td>�ړ�����w�肵�Ă�������</td>
	</tr>
	<tr>
		<td></td>
		<td><FONT COLOR="4169E1"><SMALL>�i�ߑO�A�ߌ�A�[�ϗ\����\�ł��j</SMALL></FONT></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<SELECT NAME="SELECT">
<%
	For i = 0 To iTimeCnt - 1
%>
				<OPTION VALUE=<%=TimeSlot(i)%> ><%=TimeName(i)%>
<%
	Next
%>
			</SELECT>
			<input type="submit" value="�@���s�@" id=submit2 name=submit2>
		</td>
	</tr>
</table>
</form>
</center>

<center>
<form  METHOD="post" NAME="CHG" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=CHG&OPENO=<%=sOpeNo%>" onSubmit="return ClickChg()">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">����</td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td>��������̗\��ԍ����w�肵�Ă��������B</td>
	</tr>
	<tr>
		<td></td>
		<td><FONT COLOR="4169E1"><SMALL>�i�ߑO�A�ߌ�A�[�ϗ\��̑�����w��ł��܂��j</SMALL></FONT></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<input type="text" NAME="CHGOPE" SIZE="9" MAXLENGTH="5">
			<input type="submit" value="�@���s�@" id=submit3 name=submit3>
		</td>
	</tr>
</table>
</form>
</center>

</table>
</center>

<center>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="�@���~�@" id=submit6 name=submit6></td>
	</form>
</center>

</body>
</html>
