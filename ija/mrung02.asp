<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim vCtnoS, vCtnoE, vUserID
Dim sCntNo,sCntNo2
Dim sUserID
Dim sSQL
Dim sErrMsg
Dim sErrOpt

sErrMSg = ""
sErrOpt = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

vCtnoE = Trim(Request.QueryString("cont_e"))
vCtnoS = Trim(Request.QueryString("cont_s"))
vUserID = Trim(Request.QueryString("UserID"))

If (IsEmpty(vCtnoE) Or vCtnoE = "") And (IsEmpty(vCtnoS) Or vCtnoS = "") Then
	sErrMsg = "�R���e�i������"
Else
	If IsEmpty(vUserID) Or vUserID = "" Then
		sErrMsg = "���[�U�[ID������"
	Else
		sUserID = vUserID
	End If
End If

If sErrMsg = "" Then
	Dim conn, rs
	ConnectSvr conn, rs

	'�Y������R���e�i��T��
	If IsEmpty(vCtnoE) Or vCtnoE = "" Then
		'�R���e�i�ԍ��̐��l�����̂ݓ��͂���Ă���ꍇ
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container GROUP BY RTrim([ContNo]), ContNo "
		sSQL = sSQL & "HAVING (((RTrim([ContNo])) Like '%" & vCtnoS & "'))"
	Else
		'�R���e�i�ԍ��̉p�������A���l�����Ƃ��ɓ��͂���Ă���ꍇ
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container "
		sSQL = sSQL & "WHERE RTrim([ContNo]) = '" & UCase(vCtnoE) & vCtnoS & "'"
	End If
	rs.Open sSQL, conn, 0, 1, 1
	If rs.Eof Then
		sErrMsg = "�Y���R���e�i�Ȃ�"
		sErrOpt = vCtnoE & vCtnoS
	Else
		sCntNo = rs("CT")		'�R���e�i�ԍ��Đݒ�
		rs.MoveNext
		Do While Not rs.EOF
			sCntNo2 = rs("CT")
			rs.MoveNext
			If sCntNo<>sCntNo2 Then
				sErrMsg = "���ŕ�������"
				sErrOpt = vCtnoS
				Exit Do
			End If
		Loop
	End If
	rs.Close

	If sErrMsg = "" Then
		' ���񌟍������R���e�i�ԍ������[�U�e�[�u���ɕۑ�(����Ƀf�t�H���g�ŕ\�������)
		sSQL = "SELECT lUserTable.BeforeCntnrNo FROM lUserTable WHERE lUserTable.UserID='" & sUserID & "'"
		rs.Open sSQL, conn, 2, 2
		If Not rs.Eof Then
			rs("BeforeCntnrNo") = sCntNo
			rs.Update
		End If
		rs.Close
	End If

	conn.Close
End If

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If sErrMsg<>"" Then
	WriteLogM oFS, sUserID, "6201", "�g��-���������R���e�i�ԍ�����", "10",sPhoneType, vCtnoE & "/" & vCtnoS & "," & "���͓��e�̐���:1(���)" & sErrMsg
Else
	WriteLogM oFS, sUserID, "6201", "�g��-���������R���e�i�ԍ�����", "10",sPhoneType, vCtnoE & "/" & vCtnoS & "," & "���͓��e�̐���:0(������)"
	WriteLogM oFS, sUserID, "6202", "�g��-������������", "00",sPhoneType, sCntNo & ","
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="�^�s������">
		<center>
		�y�����������́z<br>
<%
		If sErrMsg <> "" Then
%>
			<center>
			<%=sErrOpt%><br>
			<center>
			<%=sErrMsg%><br><br>
			<center>
			<a task="gosub" dest="index.asp">�ƭ�</a>
<%
		Else
%>
			<center>
			�R���e�i�ԍ�<br>
			<center>
			<%=sCntNo%><br>
			<center>
			�����������<br>
			<center>
			<a task="gosub" accesskey="3"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=C">�o:��q�ɒ�
			</a><br>
			<center>
			<a task="gosub" accesskey="4"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=D">�o:����ݸފ�
			</a><br>
			<center>
			<a task="gosub" accesskey="1"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=A">��:�����q�ɒ�
			</a><br>
			<center>
			<a task="gosub" accesskey="2"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=B">��:�f�o����
			</a><br>
<%
		End If
%>
	</display>
	</hdml>
<%
Else
	' EzWeb�ȊO�̃^�O��ҏW
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("�^�s������")%>
	</head>
	
	<body>
	<center>
	�y�����������́z
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<%=sErrOpt%><br>
		<%=sErrMsg%><br><br>
		<form action="index.asp" method="get">
			<input type="submit" value="�ƭ�">
		</form>
<%
	Else
%>
		<form action="mrung03.asp" method="get">
			�R���e�i�ԍ�<br>
			<%=sCntNo%><br>

			�����������<br>
			<select name="operation">
				<option value="C">�o:��q�ɒ�</option>
				<option value="D">�o:����ݸފ�</option>
				<option value="A">��:�����q�ɒ�</option>
				<option value="B">��:�f�o����</option>
			</select>
			<br><br>
			<input type="hidden" name="ContNo" value="<%=sCntNo%>">
			<input type="hidden" name="UserID" value="<%=sUserID%>">
			<input type="submit" value="����">
		</form>
<%
	End If
%>
	<hr>
	</body>
	</html>
<%
End If
%>
