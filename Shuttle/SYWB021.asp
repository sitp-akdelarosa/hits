
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
/* ����{�^�� */
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
}

/* �ڍ׊m�F�{�^�� */
function ClickSend1(go) {

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

/*�����*/
	if (document.UPLOAD1.sy_zaiko.value != "")  {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.sy_zaiko.value.toUpperCase()
		return true;
	}

/*�r�x�݌ɑI�����*/
	if (document.UPLOAD1.SELECT1.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.SELECT1.value
		return true;
	}

/*�r�x��݌ɑI�����*/
	if (document.UPLOAD1.SELECT2.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.SELECT2.value
		return true;
	}
}
//--->
</SCRIPT>

</head>

<body>
<%
	Dim conn, rsd, sql											'�c�a�ڑ�
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator			'���[�U���
	Dim sYMD													'�w����t
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

	'�V���[�V�\�����̎擾
	sDispChassis1 = ""
	sDispChassis2 = ""
	sChk1 = ""
	sChk2 = ""
	
	sPlateNo = sDispChassis1 & "�@" & sDispChassis2

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title26.gif" width="236" height="34"><p>
</center>

		<font face="�l�r �S�V�b�N">
   
<center>
<form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB024.asp?YMD=<%=sYMD%>" onSubmit="return ClickSend()">
<table border="1" width="420"  >
<b><font color=#000080>�ΏۃV���[�V</font></b>
		<tr bgcolor=#ffff99><td><br>
				�r�x�݌ɂ��I���@�@<SELECT NAME="SELECT1">
						<OPTION VALUE="No0" >�@
						<%	i = 1
							sql = "SELECT * FROM sChassis" & _
								  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
								  "  AND StackFlag <> ' '"
							sql = sql & "  Order By ChassisId"
							rsd.Open sql, conn, 0, 1, 1
			
							if not rsd.eof then
								do while not rsd.EOF%>
									<OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
									<%rsd.MoveNext
									i = i + 1
								loop
							end if
							rsd.Close
						%>
					</SELECT><br>
				�r�x��݌ɂ��I���@<SELECT NAME="SELECT2">
						<OPTION VALUE="No0" >�@
						<%	i = 1
							sql = "SELECT * FROM sChassis" & _
								  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
								  "  AND StackFlag = ' '"
							sql = sql & "  Order By ChassisId"
							rsd.Open sql, conn, 0, 1, 1
							if not rsd.eof then
								do while not rsd.EOF%> 
									<OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
									<%rsd.MoveNext
									i = i + 1
								loop
							end if
							rsd.Close
						%>
					</SELECT><br>
	����͂���ꍇ�@�@�@<INPUT TYPE="text" NAME="sy_zaiko" SIZE="9" MAXLENGTH="5" value=<%=sDispChassis1%>>�@<input type="button" value="�V���[�V�̏ڍ׊m�F" id=submit5 name=submit5 onclick="return ClickSend1(this)">
		<br>
		</td></tr>
</table><br><br>
</center>
<br>
<center>
<table border=0>
		<td><input type="submit"  value="�@����@" id=submit4 name=submit4></td>
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