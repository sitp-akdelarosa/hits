<!--<%@ LANGUAGE="VBScript" %>-->
<%
'Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="Sywb060.inc"-->
<html>

<head>
<title>��o���\�񌋉ʉ��</title>
</head>

<body>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
		<p align="center"><img border="0" src="image/title34.gif" width="236" height="34"><p>
<%
	Dim sYMD, Idx
	Dim conn, rsd
	Dim sName, sErrMsg
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sVPBook, sVSL, sVOY, sLINE, sTERM, sSIZE, sTYPE
	Dim sHIGHT, sMATERIAL
	Dim sTERM_Name, sTYPE_Name, sMATERIAL_Name
	Dim sContType_M(14), sTerminal_M(49), sMaterial_M(9)
	Dim sOpeNo
	Dim sDeliverTo

	sOpeNo = "10023"
	'�w����t�擾(����ʈ��p)
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD, 9, 2)
	sYMD = Left(sYMD, 8)
	'��Ǝ��ԑ�(����ʈ��p)
	sName = TRIM(Request.QueryString("NAME"))
	'�I���u�b�L���O�ԍ�
	sVPBook = TRIM(Request.QueryString("BOOK"))
	'�I��{�D
	sVSL = TRIM(Request.QueryString("VSL"))
	'�I�����q
	sVOY = TRIM(Request.QueryString("VOY"))
	'�I��D��
	sLINE = TRIM(Request.QueryString("LINE"))
	'�I���^�[�~�i��
	sTERM = TRIM(Request.QueryString("TERM"))
	'�I���T�C�Y
	sSIZE = TRIM(Request.QueryString("SIZE"))
	'�I���^�C�v
	sTYPE = TRIM(Request.QueryString("TYPE"))
	'�I���n�C�g
	sHIGHT = TRIM(Request.QueryString("HIGHT"))
	'�I���ގ�
	sMATERIAL = TRIM(Request.QueryString("MATERIAL"))
	'�R���e�i���o��
	sDeliverTo = TRIM(Request.QueryString("DELIVERTO"))
%>		<center>
		<table border="1">   
			<tr ALIGN=middle>
				<td width="120" bgcolor ="#e8ffe8">��Ǝ���</td>
				<td width="360" ><%=ChgYMDStr2(sYMD)%>�@<%=sName%></td>
			</tr>
		</table>
		<br>
<%
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

	'�^�[�~�i�����̂̎擾
	sTERM_Name =  GetTerminal_Name(conn, rsd, sTERM)
	'�R���e�i�^�C�v���̂̎擾
	sTYPE_Name =  GetContType_Name(conn, rsd, sTYPE)
	'�R���e�i�ގ����̂̎擾
	sMATERIAL_Name =  GetMaterial_Name(conn, rsd, sMATERIAL)
%>
	<table border="1" width="500"  >
		<tr><td width="120" bgcolor="#cccc99">�u�b�L���O�ԍ�</td>
		<td><%=sVPBook%></td>
		</tr>
		<tr>
			<td width="120" bgcolor="#cccc99">�Ώۃo���v�[��</td>
			<td><%=sTERM_Name%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">�T�C�Y</td>
			<td><%=sSIZE%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">�^�C�v</td>
			<td><%=sTYPE_Name%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">����</td>
			<td><%=sHIGHT%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">�ގ�</td>
			<td><%=sMATERIAL_Name%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">�{�D</td>
			<td><%=sVSL%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">���q</td>
			<td><%=sVOY%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">�D��</td>
			<td><%=sLINE%></td>
		</tr>
	</table><br>
<%

	'��o���\��X�V����
	Call UpdateApp_VPBook(conn, rsd, sUsrID, sGrpID, _
			sYMD, sHH, sVSL, sVOY, sLINE, sVPBook, _
            sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
            sErrMsg, sOpeNoVan) 

'�f�[�^���Ȃ��ꍇ�ق��̃G���[
	if sErrMsg <> "" then
%>		<center>	<%
			Response.Write "<center><FONT color=Red><U>" & "�i���ʁj�F�s�@" & sErrMsg & "</U></FONT></center>"
			%><br>
		</center>

<%	Else	%>
		<center><FONT  size=4 color=blue><U>�i���ʁj�F�n�j�@�\��ԍ��F<%=sOpeNoVan%></FONT></center><br></U>

		<center>
			<table>
			<tr>
			<td><font color=red>�i���Ӂj<U>�{���s���̗��R�ŗ\��s�ɂȂ�\��������܂��B�P�O����ȍ~�Ɉꗗ��</U></font><br></td>
			<tr>
			<td>�@�@�@�@ <font color=red><u>�ʂ��Ċm�F���Ă��������܂��悤���肢�v���܂��B</u></font><br></td>
			</tr>
			<tr>
			<td>�@�@�@�@<font color=red>�i���ڈ������Ƃ̏d�����������ꍇ�j</font><br></td></tr>
			</table>
		</center>

<%
	End If 
	conn.Close
%>
	<br>
	<center>
	<table border=0>
	    <form id=form1 name=form1>
	    <td><input type="button" value="�@�߂�@" onclick="history.back()"  id=button1 name=button1></td>
		</form>

	    <form  METHOD="post"  NAME="BACK" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="�ꗗ��ʂ�" id=submit2 name=submit2></td>
		</form>
	</table>
	</center>

</body>
</html>
