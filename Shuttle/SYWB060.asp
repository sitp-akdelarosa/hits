<!--<%@ LANGUAGE="VBScript" %>-->
<%
'Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="Sywb060.inc"-->
<html>
<head>
<title>��o���\����(�u�o�u�b�L���O)</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
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
<%
	Dim sYMD, sYMD_OLD, Idx
	Dim conn, rsd
	Dim sName
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sVPBook, sErrMsg, sVPLast

	Dim sContType_M(14), sTerminal_M(49), sMaterial_M(9)

	'�w����t�擾(����ʈ��p)
	sYMD_OLD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD_OLD, 9, 2)
	sYMD = Left(sYMD_OLD, 8)
	'��Ǝ��ԑ�(����ʈ��p)
	sName = TRIM(Request.QueryString("NAME"))
	'�I���b�x�^�u�o
	sVPBook = TRIM(Request.Form("VPBookNo"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	If sGrpID = "" Then
		Response.Write "���[�U���o�^����Ă��܂���B(" & sUsrID & ")"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.End
	End If

	'01/12/05 ��o���L�ςݎw��L���擾
	sVPLast = GetEnv(conn, rsd, "VPLastFlag")

	If sVPLast = "N" and sName = "�[�ώw��" Then
		sErrMsg = "��o���\��̗[�ώw��͂ł��܂���"
	Else

		'���[�U���̎擾
		Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

		'�R���e�i�^�C�v�擾
		Call GetContType(conn, rsd, sContType_M)

		'�R���e�i�ގ��擾
		Call GetMaterial(conn, rsd, sMaterial_M)

		'�Ώۂu�o�擾
		Call GetTerminal(conn, rsd, sTerminal_M)

		'�u�b�L���O�\��`�F�b�N�����P�i��Ɠ��̃`�F�b�N�j
		Call VPBookCheck1(conn, rsd, sUsrID, sGrpID, _
				sYMD, sHH, sVPBook,	sErrMsg)

		If sErrMsg = "" Then
		'�u�o�u�b�L���O���R�[�h���̓ǂݍ��݂��s��
			Call GetVPBooking1(conn, rsd, sVPBook, sErrMsg)		
		End If
	End If

	If sErrMsg = "" then
		Idx = 1
		rsd.MoveFirst
		Do Until rsd.EOF
%>
			function ChkGo<%=Idx%>() {
				deliverto=document.form0.DeliverTo.value
				if ( !ChkChara(deliverto) ) {
					window.alert("�R���e�i���o��͔��p���[�}���œ��͂��Ă��������B");
					return;
				}
				str="SYWB061.asp?YMD=<%=sYMD_OLD%>&NAME=<%=sName%>&VSL=<%=trim(rsd("VslCode"))%>&VOY=<%=trim(rsd("Voyage"))%>&LINE=<%=trim(rsd("LineCode"))%>&BOOK=<%=trim(rsd("BookNo"))%>&TERM=<%=trim(rsd("Terminal"))%>&SIZE=<%=trim(rsd("ContSize"))%>&TYPE=<%=trim(rsd("ContType"))%>&HIGHT=<%=trim(rsd("ContHeight"))%>&MATERIAL=<%=trim(rsd("Material"))%>"
				if ( confirm('�\�񂵂܂����H') )
				{
					location.href=str + "&DELIVERTO=" + deliverto;
				}
			}
<%
			rsd.MoveNext
			Idx = Idx + 1
		Loop
		rsd.MoveFirst
	End If
%>
</SCRIPT>
</head>

<body>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>

<%
'�f�[�^���Ȃ��ꍇ�̏���

	if sErrMsg <>  "" then
%>		<center>
			<p><img border="0" src="image/title34.gif" width="236" height="34"><p>
			<table border="1">   
				<tr ALIGN=middle>
					<td width="120" bgcolor ="#e8ffe8">��Ǝ���</td>
					<td width="360" ><%=ChgYMDStr2(sYMD)%>�@<%=sName%></td>
				</tr>
			</table>
			<br>
			<table border="1" width="500"  >
				<tr><td width="160" bgcolor="#cccc99">�u�b�L���O�ԍ�</td>
					<td><%=sVPBook%></td>
				</tr>
			</table><br><%
			Response.Write "<center><FONT color=Red><U>" & "�i���ʁj�F�s�@" & sErrMsg & "</U></FONT></center>"
			%><br>
		</center>

		<br>     
		<center>
			<table border=0>
			    <form id=form1 name=form1>
			    <input type="button" value="�@�߂�@" onclick="history.back()"  id=button1 name=button1>
				</form>
			</table>
		</center>
<%	Else	%>
		<center>
		<p><img border="0" src="image/title33.gif" width="236" height="34"><p>
		<table border="0">   
			<tr ALIGN=middle>
				<td><font size=5><u><%=ChgYMDStr2(sYMD)%>�@<%=sName%>�@��o���֗\��</u></font></td>
			</tr>
			<tr></tr><tr></tr><tr></tr><tr></tr><tr></tr>
			<tr ALIGN=middle>
				<th><font size=4>�u�b�L���O�ԍ��E�E�E<%=trim(rsd("BookNo"))%></font></th>
			</tr>
			<tr></tr><tr></tr><tr></tr>
		<table border="0" bgcolor ="#FFFFBB" width="420">   
			<tr ALIGN=middle>
				<td>�R���e�i���o�����͌�ړI�̂��̂�I�����Ă�������</td>
			</tr>
			<tr ALIGN=middle>
			    <form id=form0 name=form0>
				<td>
				�R���e�i���o��F�@
				<INPUT NAME="DeliverTo" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
				(���p���[�}���œ��͂��Ă�������)</td>
				</form>
			</tr>
		</table>

		</table>
		<br>

		<table border="1">   
			<tr ALIGN=middle>
				<td width="50" BGCOLOR=#7FFFD4></td>
				<th width="150" BGCOLOR=#7FFFD4>�Ώۃo���v�[��</th>
				<th width="50" BGCOLOR=#7FFFD4>�T�C�Y</th>
				<th width="100" BGCOLOR=#7FFFD4>�^�C�v</th>
				<th width="50" BGCOLOR=#7FFFD4>����</th>
				<th width="100" BGCOLOR=#7FFFD4>�ގ�</th>
				<th width="50" BGCOLOR=#7FFFD4>�{�D</th>
				<th width="50" BGCOLOR=#7FFFD4>���q</th>
				<th width="50" BGCOLOR=#7FFFD4>�D��</th>
			</tr>

<%
			Idx = 1
			rsd.MoveFirst
			Do Until rsd.EOF	%>
					<tr ALIGN=middle>
						<td><font size=4>
<!---						<A href="SYWB061.asp?YMD=<%=sYMD_OLD%>&
												NAME=<%=sName%>&
												VSL=<%=trim(rsd("VslCode"))%>&
												VOY=<%=trim(rsd("Voyage"))%>&
												LINE=<%=trim(rsd("LineCode"))%>&
												BOOK=<%=trim(rsd("BookNo"))%>&
												TERM=<%=trim(rsd("Terminal"))%>&
												SIZE=<%=trim(rsd("ContSize"))%>&
												TYPE=<%=trim(rsd("ContType"))%>&
												HIGHT=<%=trim(rsd("ContHeight"))%>&
												MATERIAL=<%=trim(rsd("Material"))%>" onclick="JavaScript:return confirm('�\�񂵂܂����H')"><%=Idx%></a>
--->
<A href="JavaScript:ChkGo<%=Idx%>();"><%=Idx%></a>

						</font></td>
						<td><%=SetTerminal(rsd("Terminal"), sTerminal_M)%></td>
						<td><%=rsd("ContSize")%></td>
						<td><%=SetContType(rsd("ContType"), sContType_M)%></td>
						<td><%=rsd("ContHeight")%></td>
						<td><%=SetMaterial(rsd("Material"), sMaterial_M)%></td>
						<td><%=rsd("VslCode")%></td>
						<td><%=rsd("Voyage")%></td>
						<td><%=rsd("LineCode")%></td>
					</tr>
<%				rsd.MoveNext
				Idx = Idx + 1
			Loop
%>				
		</table>
		<br>
		<center>
			�ړI�̂��̂�I�����Ă��������i���̔ԍ����N���b�N���ĉ������j
		    <form id=form1 name=form1>
		    <input type="button" value="�@���~�@" onclick="history.back()"  id=button1 name=button1>
			</form>
		</center>

<%
	End If 
	conn.Close
%>

</body>
</html>
