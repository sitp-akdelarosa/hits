<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common_cam.inc"-->
<!--#include file="mcommon.inc"-->
<%
Const PAGE_SIZE = 10
Dim vPageNo
Dim nPageNo
Dim nPageMax
Dim vBlno
Dim sBlno
Dim sql
Dim sErrMsg
Dim sErrOpt
Dim iRecCnt
Dim sContNo(20)
Dim nContCount
Dim nContPtr

sErrMsg = ""
sErrOpt = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

vPageNo = Request.QueryString("PAGENO")
vBlno = Request.QueryString("BLno")

If IsEmpty(vPageNo) Or vPageNo ="" Then
	vPageNo = "1"
	vBlno = Ucase(Trim(Request.QueryString("BLno")))
End If
nPageNo = CInt(vPageNo)

Dim conn, rs
ConnectSvr conn, rs

iRecCnt = 0
If vBlno = "" Then
	sErrMsg = "�a�^�k������"
Else
	sBlno = Trim(vBlno)

	' �a�k�ԍ�like����
	If Len(sBlno) <= 19 Then
		dim iblcnt
		dim slblno

		iblcnt = 0
		slblno = "%" & sBlno
		sql = "SELECT RTrim([BLNo]) AS BLN FROM ImportCont GROUP BY RTrim([BLNo]), BLNo "
		sql = sql  & "HAVING (((RTrim([BLNo])) Like '" & slblno & "'))"
		rs.Open sql, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "�Y���a�^�k�Ȃ�"
			sErrOpt = Trim(vBlno)
		Else
			sBlno = rs("BLN")		'�a�k�ԍ��Đݒ�
			rs.MoveNext
			If Not rs.Eof Then
				sErrMsg = "BL�������݂��܂�"
				sErrOpt = Trim(vBlno)
			End If
		End If
		rs.Close
	End If
End If

If sErrMsg = "" Then
'--- mod by MES(2004/9/10)
'	sql = "SELECT ContNo, FreeTime, OLTICFlag, OLTICNo, CYDelTime, " & _
'		" DOStatus, DelPermitDate, OLTDateFrom, OLTDateTo, FreeTimeExt " & _
'		" FROM ImportCont WHERE BLNo='" & sBlno & "' " & _
'		" ORDER BY ContNo"
'--- mod by MES(2005/3/28)
'	sql = "SELECT ContNo, FreeTime, OLTICFlag, OLTICNo, OLTICDate, CYDelTime, " & _
'		" DOStatus, DelPermitDate, OLTDateFrom, OLTDateTo, FreeTimeExt " & _
'		" FROM ImportCont WHERE BLNo='" & sBlno & "' " & _
'		" ORDER BY ContNo"
	sql = "SELECT ImportCont.ContNo, ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, " & _
		" ImportCont.OLTICDate, ImportCont.CYDelTime, ImportCont.DOStatus, " & _
		" ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, ImportCont.FreeTimeExt, " & _
		" Container.ListNo, Container.OffDockFlag, Container.DsListFlg " & _
		" FROM ImportCont, Container WHERE ImportCont.BLNo='" & sBlno & "' " & _
		" AND Container.ContNo=ImportCont.ContNo AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
		" ORDER BY Container.ContNo"
'--- end MES
'--- end MES
	rs.Open sql, conn, 0, 1, 1
	If rs.eof Then
		sErrMsg = "�Y���a�^�k�Ȃ�"
		sErrOpt = Trim(vBlno)
	Else
		nContCount = 0
		Do While Not rs.Eof
			'���o�\
			If CanCarryOut(rs)="Y" Then
				iRecCnt = iRecCnt + 1							'���ۃf�[�^�J�E���g
				If (nPageNo - 1) * PAGE_SIZE < iRecCnt And iRecCnt <= nPageNo * PAGE_SIZE Then
					nContCount = nContCount + 1
					sContNo(nContCount) = Trim(rs("ContNo"))
				End If
			End If
			rs.MoveNext
		loop

		If iRecCnt =  0 Then
			sErrOpt = "���o�\"
			sErrMsg = "�R���e�i����"
		Else
			'�S�y�[�W��
			nPageMax = -Int(-iRecCnt / PAGE_SIZE)
		End If
	End If
	rs.Close
End If
conn.Close

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If sErrMsg<>"" Then
	WriteLogM oFS, "Unknown", "2403", "�g��-BL�ԍ��Ɖ�i�����ӓ��j", "10",sPhoneType, sBLNo & "," & "���͓��e�̐���:1(���)" & sErrMsg
Else
	WriteLogM oFS, "Unknown", "2403", "�g��-BL�ԍ��Ɖ�i�����ӓ��j", "10",sPhoneType, sBLNo & "," & "���͓��e�̐���:0(������)"
	WriteLogM oFS, "Unknown", "2404", "�g��-�R���e�i�ԍ��ꗗ�i�����ӓ��j", "00",sPhoneType, nPageNo & "/" & nPageMax & ","
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="BL�ԍ��Ɖ�">
	<center>
	�yBL�ԍ��Ɖ�z<br>
<%
	If sErrMsg <> "" Then
%>
		<center>
		<%=sErrOpt%><br>
		<center>
		<%=sErrMsg%><br><br>
<%
	Else
		For nContPtr = 1 To nContCount
%>
			<center>
			<a task="gosub" accesskey="<%=CStr(nContPtr Mod 10)%>"
						dest="mcont02cam.asp?Ctno=<%=sContNo(nContPtr)%>">
				<%=sContNo(nContPtr)%>
			</a><br>
<%
		Next
%>
		<center>
		-<%=nPageNo%>/<%=nPageMax%>-<br>
		<center>
<%
		If 1 < nPageNo Then
%>
			<a task="gosub" dest="mblno02cam.asp?PAGENO=<%=nPageNo - 1%>&BLno=<%=sBlno%>">�O��</a>
<%
		End If
		If nPageNo < nPageMax Then
%>
			<a task="gosub" dest="mblno02cam.asp?PAGENO=<%=nPageNo + 1%>&BLno=<%=sBlno%>">����</a>
<%
		End If
%>
		<br>
<%
	End If
%>
	<center>
	<a task="gosub" dest="index.asp">�ƭ�</a>
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
		<%=GetTitleTag("BL�ԍ��Ɖ�")%>
	</head>
	<body>
	<center>
	�yBL�ԍ��Ɖ�z
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<%=sErrOpt%><br>
		<%=sErrMsg%><br><br>
<%
	Else
		For nContPtr = 1 To nContCount
%>
			<a href="mcont02cam.asp?Ctno=<%=sContNo(nContPtr)%>" <%=GetKeyTag(CStr(nContPtr))%>>
				<%=GetKeyLabel(CStr(nContPtr))%><%=sContNo(nContPtr)%>
			</a><br>
<%
		Next
%>
		-<%=nPageNo%>/<%=nPageMax%>-<br>
<%
		If 1 < nPageNo Then
%>
			<a href="mblno02cam.asp?PAGENO=<%=nPageNo - 1%>&BLno=<%=sBlno%>">�O��</a>
<%
		End If
		If nPageNo < nPageMax Then
%>
			<a href="mblno02cam.asp?PAGENO=<%=nPageNo + 1%>&BLno=<%=sBlno%>">����</a>
<%
		End If
%>
		<br>
<%
	End If
%>
	<form action="../index.asp" method="get">
		<input type="submit" value="�ƭ�">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
