
<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>�V���[�V�����ݒ���</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
var  f1=false;
var  f2=false;
var  f3=false;

/* �o�^�{�^�� */
function ClickSend() {

	/* ���̓`�F�b�N */

	if  (document.UPLOAD1.SELECT1.value == "No0" && document.UPLOAD1.SELECT2.value == "No0" &&
		document.UPLOAD1.sy_zaiko.value == "") {
		window.alert("�V���[�V�h�c�ɖ���������܂��B");
		return false;
	}
	if ((document.UPLOAD1.sy_zaiko.value != ""   && document.UPLOAD1.SELECT1.value != "No0") || 
		(document.UPLOAD1.sy_zaiko.value != ""   && document.UPLOAD1.SELECT2.value != "No0") ||
		(document.UPLOAD1.SELECT1.value != "No0" && document.UPLOAD1.SELECT2.value != "No0")) {
		window.alert("�V���[�V�h�c�ɖ���������܂��B");
		return false;
	}


	/* �����I�����[�j���O�`�F�b�N */
	/*if  ((f1==true) && (f2==true)) {
		window.alert("�����̐ݒ�ɖ���������܂��B");
		return false;
	}*/
	if  ((document.UPLOAD1.check1.checked==true) && (document.UPLOAD1.check2.checked==true)) {
		window.alert("�����̐ݒ�ɖ���������܂��B");
		return false;
	}
}
/* �ڍ׊m�F�{�^�� */
function ClickSend2(go) {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.sy_zaiko.value.toUpperCase()
		return true;
}
//--->
</SCRIPT>

</head>

<body>
<%
	Dim conn, rsd, sql											'�c�a�ڑ�
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator			'���[�U���
	Dim sYMD, sChassisID										'�w����t�A�V���[�VID
	Dim sDispChassis1, sDispChassis2, sPlateNo, sChk1, sChk2	'�V���[�V�\�����
	Dim i
	Dim sNO, sChkChassisID										'���g�p
	Dim sErr1, sErr2, sErr3, sChassis							'���g�p

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))

	'�V���[�VID�擾
	sChassisID = TRIM(Request.QueryString("TRGID"))
	If sChassisID = "" Then
		If Request.Form("sy_zaiko") <> "" Then
			sChassisID = Request.Form("sy_zaiko")	'�����
		ElseIf Request.Form("SELECT1")  <> "No0" Then
			sChassisID = Request.Form("SELECT1")		'�݌ɑI��
		Else
			sChassisID = Request.Form("SELECT2")		'��݌ɑI��
		End If
	End If

	'�V���[�V�\�����̎擾
	sDispChassis1 = ""
	sDispChassis2 = ""
	sChk1 = ""
	sChk2 = ""
	sql = "SELECT * FROM sChassis" & _
			" WHERE ChassisId = '" & sChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		sDispChassis1 = sChassisID				'�w��V���[�VID
		sDispChassis2 = trim(rsd("PlateNo"))	'�v���[�g�ԍ�
		If rsd("NotDelFlag") = "Y" Then
			sChk1 = "1"		'���o���ڂ��Ȃ��V���[�V
		End If
		If rsd("NightFlag") = "Y" Then
			sChk2 = "1"		'�[�ς݃V���[�V
		End If
	End If
	rsd.Close
	
	sPlateNo = sDispChassis1 & "�@" & sDispChassis2

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title26.gif" width="236" height="34"><p>
</center>

		<font face="�l�r �S�V�b�N">
   
<center>
<form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB032.asp?TDATE=<%=sYMD%>" onSubmit="return ClickSend()">
<table border="1" width="420"  >
<b><font color=#000080>�ΏۃV���[�V</font></b>
		<tr bgcolor=#ffff99><td><br>�@<%=sPlateNo%>
		<br>�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@<input type="button" value="�V���[�V�̏ڍ׊m�F" id=submit5 name=submit5 onclick="return ClickSend2(this)">
		<br>		<INPUT TYPE=hidden NAME="sy_zaiko" VALUE=<%=sDispChassis1%>>
					<INPUT TYPE=hidden NAME="SELECT1" VALUE="No0">
					<INPUT TYPE=hidden NAME="SELECT2" VALUE="No0">
		</td></tr>
</table><br><br>
<table border="1" width="420">
<b><font color=#000080>�����I��</font></b>	<tr bgcolor=#ccffcc><td><br>
				<%	if sChk1 = "1" then %>
<INPUT TYPE=checkbox NAME="check1" checked onClick="f1=!f1">���o�R���e�i���ڂ��Ȃ�<br>
				<%	else	%>
<INPUT TYPE=checkbox NAME="check1" onClick="f1=!f1">���o�R���e�i���ڂ��Ȃ�<br>
				<%	end if
					if sChk2 = "1" then%>
<INPUT TYPE=checkbox NAME="check2" checked onClick="f2=!f2">�[�ς̂ݍڂ���<br>
				<%	else	%>
<INPUT TYPE=checkbox NAME="check2" onClick="f2=!f2">�[�ς̂ݍڂ���<br>
				<%	end if	%>
<INPUT TYPE=checkbox NAME="check3" onClick="f3=!f3">�O���[�v�ύX
			<SELECT NAME="SELECT3">
				<%	sql = "SELECT * FROM sMGroup" & _
						  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
					rsd.Open sql, conn, 0, 1, 1
					do while not rsd.EOF
						%><OPTION VALUE=<%=RTRIM(rsd("GroupID"))%>><%=rsd("GroupName")%>
						<%rsd.MoveNext
					loop
					rsd.Close
'�ق��̃O���[�v
					sql = "SELECT * FROM sMGroup" & _
						  " WHERE RTRIM(GroupID) <> '" & sGrpID & "'"
					rsd.Open sql, conn, 0, 1, 1
					do while not rsd.EOF
						%><OPTION VALUE=<%=RTRIM(rsd("GroupID"))%>><%=rsd("GroupName")%>
						<%rsd.MoveNext
					loop
					rsd.Close %>
			</SELECT><br><br>
			<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
	</td></tr>
</table>
</center>
<br>
<center>
<table border=0>
		<td><input type="submit"  value="�@���s�@" id=submit4 name=submit4></td>
	</form>
	<td>�@</td>
	<td>�@</td>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="�@���~�@" id=submit6 name=submit6></td>
	</form>
</table>
</center>

<br>     
<br>     
</body>     
</html>