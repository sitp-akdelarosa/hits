<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB031A.inc"-->
<!--#include file="SYWB031B.inc"-->
<!--#include file="SYWB031C.inc"-->
<!--#include file="SYWB031D.inc"-->
<html>

<head>
<title>���o���\��ύX���ʉ��</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
/* �߂�̃N���b�N */
function ClickBack() {
	history.back()
}

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sOpeNo, sCmd, sHH, sChgOpeNo, sCmdName, sDelFlag, sStatus
	Dim sRecDel, sSend, sVPLast
	Dim conn, rsd
	Dim sErrMsg

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))

	'��Ɣԍ��擾
	sOpeNo = TRIM(Request.QueryString("OPENO"))

	'���s�R�}���h�擾
	sCmd = TRIM(Request.QueryString("CMD"))
	If     sCmd = "DEL" Then	'�폜
		sCmdName = "�폜"
	ElseIf sCmd = "MOV" Then	'�ړ�
		sCmdName = "�ړ�"
	ElseIf sCmd = "CHG" Then	'����
		sCmdName = "����"
	Else						'�ύX
		If TRIM(Request.Form("DeliverTo")) <> "" Then
			sCmdName = "���o��ύX"
		Else
			sCmdName = "�������ύX"
		End If
	End If

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'�\�����̎擾�i�w���Ɣԍ��j
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	sHH      = Trim(rsd("Term"))
	sDelFlag = Trim(rsd("DelFlag"))
	sStatus  = Trim(rsd("Status"))
	rsd.Close

	'01/12/05 ��o���L�ςݎw��L���擾
	sVPLast = GetEnv(conn, rsd, "VPLastFlag")

	'�\����A�\�񎞊ԑт��L�����`�F�b�N
	sErrMsg = ""
	If sDelFlag <> "Y" And sStatus <> "03" Then
		Call CheckAppWorkDate(conn, rsd, sYMD, sHH, sErrMsg) 
	End If
	sHH = ""
	If sErrMsg = "" Then

		sChgOpeNo = ""
		sErrMsg = ""

		If     sCmd = "DEL" Then	'�폜
			Call UpdOpeDel(conn, rsd, sOpeNo, sErrMsg)
		ElseIf sCmd = "MOV" Then	'�ړ�
			sHH = TRIM(Request.Form("SELECT"))
			If sVPLast = "N"  AND sHH = "B" Then	'01/12/05
				sErrMsg = "��o���\��̗[�ώw��ւ̈ړ��͂ł��܂���"
			Else
				Call UpdOpeMov(conn, rsd, sOpeNo, sYMD, sHH, sErrMsg)
			End If
		ElseIf sCmd = "CHG" Then	'����
			sChgOpeNo = TRIM(Request.Form("CHGOPE"))
			Call UpdOpeChg(conn, rsd, sOpeNo, sYMD, sChgOpeNo, sErrMsg)
		Else						'�ύX
			If TRIM(Request.Form("DeliverTo")) <> "" Then
				sSend = TRIM(Request.Form("DeliverTo"))
			Else
				sSend = TRIM(Request.Form("ReceiveFrom"))
			End If
			Call UpdOpeUpd(conn, rsd, sOpeNo, sSend, sErrMsg)
		End If
	End If

	'�c�a�ؒf
	conn.Close
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title35.gif" width="236" height="34"><p>
</center>
<br><br>

<font face="�l�r �S�V�b�N">
<center>
	<table border="1" width="300" >
		<tr ALIGN=middle><td width="130" bgcolor ="#e8ffe8">����</td><td><%=sCmdName%></td></tr>
		<tr ALIGN=middle><td width="130" bgcolor ="#e8ffe8">�\��ԍ�</td><td><%=sOpeNo%></td></tr>
<%
	If sChgOpeNo <> "" Then
%>
		<tr ALIGN=middle><td width="130" bgcolor ="#e8ffe8">��������</td><td><%=sChgOpeNo%></td></tr>
<%
	End If
%>
	</table>

	<br><br>
	<B><U><font color=#ff0000>
<%
	If sErrMsg <> "" Then
%>
	���ʁF<%=sErrMsg%><br>
<%
	Else
%>
	���ʁFOK<br>
<%
	End If
%>
	</font></U></B>
<br><br>
<%
	If sErrMsg <> "" Then
%>
		<input type="button" value="�@�߂�@" onClick="ClickBack()">
<%
	End If
%>
</center>
<br><br>
<FORM NAME="SEND">
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
</FORM>
<%
	If sErrMsg = "" Then
%>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB013.asp?TDATE=" + document.SEND.YMD.value);
</SCRIPT>
<%
	End If
%>
</body>     
</html>     
