<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim vCtno, vCtnoE, vCtnoS
Dim CntNo,sCntNo2
Dim vFlg								'�R���e�i�Ɖ���J��("1")�^�a�k�Ɖ�J��(��)
Dim sql
Dim sErrMsg
Dim sErrOpt

sErrMSg = ""
sErrOpt = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

vCtno = Trim(Request.QueryString("Ctno"))
vCtnoE = Trim(Request.QueryString("cont_e"))
vCtnoS = Trim(Request.QueryString("cont_s"))
If IsEmpty(vCtno) Or vCtno ="" Then
	vFlg = "1"
	vCtno = Ucase(vCtnoE) & vCtnoS
End If

Dim conn, rs
ConnectSvr conn, rs

CntNo = vCtno
If CntNo = "" Then
	sErrMsg = "���Ŗ�����"
Else
	'�R���e�i�ԍ��̐��l�����̂ݓ��͂���Ă���ꍇ�A�Y������R���e�i��T��
	If vFlg = "1" And  (vCtnoE = "" Or IsEmpty(vCtnoE)) Then	
		sql = "SELECT RTrim([ContNo]) AS CT  FROM ImportCont GROUP BY RTrim([ContNo]), ContNo "
		sql = sql  & "HAVING (((RTrim([ContNo])) Like '%" & vCtnoS & "'))"
		rs.Open sql, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "�Y���R���e�i�Ȃ�"
			sErrOpt = vCtnoS
		Else
			CntNo = rs("CT")		'�R���e�i�ԍ��Đݒ�
			rs.MoveNext
			Do While Not rs.EOF
				sCntNo2 = rs("CT")
				rs.MoveNext
				If CntNo<>sCntNo2 Then
					sErrMsg = "���ŕ�������"
					sErrOpt = vCtnoS
					Exit Do
				End If
			Loop
		End If
		rs.Close
	End If
End If

If sErrMSg = "" Then
'--- mod by MES(2004/9/10)
'	sql = "SELECT ImportCont.ContNo, ImportCont.DGFlag, ImportCont.WHArSchDate, ImportCont.RFFlag, " & _
'	      " ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.CYDelTime, " & _
'	      " ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, " & _
'	      " ImportCont.FreeTimeExt, Container.ContSize, Container.ContHeight, " & _
'		  " BL.RecTerminal, BL.RFFlag BRFFlag, BL.DGFlag BDGFlag " & _
'		  " FROM ImportCont, Container, BL " & _
'		  " WHERE Container.ContNo='" & CntNo & "' " & _
'		  " AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
'	      " AND Container.ContNo=ImportCont.ContNo " & _
'	      " AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl " & _
'	      " AND BL.BLNo=*ImportCont.BLNo"
'--- mod by MES(2005/3/28)
'	sql = "SELECT ImportCont.ContNo, ImportCont.DGFlag, ImportCont.WHArSchDate, ImportCont.RFFlag, " & _
'	      " ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.OLTICDate, ImportCont.CYDelTime, " & _
'	      " ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, " & _
'	      " ImportCont.FreeTimeExt, Container.ContSize, Container.ContHeight, " & _
'		  " BL.RecTerminal, BL.RFFlag BRFFlag, BL.DGFlag BDGFlag " & _
'		  " FROM ImportCont, Container, BL " & _
'		  " WHERE Container.ContNo='" & CntNo & "' " & _
'		  " AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
'	      " AND Container.ContNo=ImportCont.ContNo " & _
'	      " AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl " & _
'	      " AND BL.BLNo=*ImportCont.BLNo"
	sql = "SELECT ImportCont.ContNo, ImportCont.DGFlag, ImportCont.WHArSchDate, ImportCont.RFFlag, " & _
	      " ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.OLTICDate, ImportCont.CYDelTime, " & _
	      " ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, " & _
	      " ImportCont.FreeTimeExt, Container.ContSize, Container.ContHeight, " & _
	      " Container.ListNo, Container.OffDockFlag, Container.DsListFlg, " & _
		  " BL.RecTerminal, BL.RFFlag BRFFlag, BL.DGFlag BDGFlag " & _
		  " FROM ImportCont, Container, BL " & _
		  " WHERE Container.ContNo='" & CntNo & "' " & _
		  " AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
	      " AND Container.ContNo=ImportCont.ContNo " & _
	      " AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl " & _
	      " AND BL.BLNo=*ImportCont.BLNo"
'--- end MES
'--- end MES
	rs.Open sql, conn, 0, 1, 1
	If rs.eof Then
		sErrMsg = "�Y���R���e�i�Ȃ�"
		sErrOpt = CntNo
	Else
		' �ꏊ�^�R���e�i�T�C�Y
		Dim sPlace
		sPlace = Trim(rs("RecTerminal")) & "�^" & Trim(rs("ContSize")) & "ft"

		' �댯��
		Dim sDanger
		sDanger=rs("DGFlag")
		If IsNull(sDanger) Or sDanger="" Then
			sDanger=rs("BDGFlag")
		End If
		If sDanger = "H" Then
			sDanger = "�댯��:��"
		Else
			sDanger = "�댯��:�|"
		End If

		' �q�ɓ����w������
		Dim sArriveTime, sYear, sMonth, sDay, sHour, sMinute
		sArriveTime = "�q�ɓ����w������<br>�@"
		If Not IsNull(rs("WHArSchDate")) Then
			sYear = CStr(Year(rs("WHArSchDate")))
			sMonth = Right(CStr(Month(rs("WHArSchDate")) + 100), 2)
			sDay = Right(CStr(Day(rs("WHArSchDate")) + 100), 2)
			sHour = Right(CStr(Hour(rs("WHArSchDate")) + 100), 2)
			sMinute = Right(CStr(Minute(rs("WHArSchDate")) + 100), 2)
			sArriveTime = sArriveTime & sYear & "/" & sMonth & "/" & sDay & "�@"  & sHour & ":" & sMinute
		End If

		' ����
		Dim sHeight
		sHeight = "����:" & Trim(rs("ContHeight"))

		' ���[�t�@�[
		Dim sReefer
		sReefer = rs("RFFlag")
		If IsNull(sReefer) Or sReefer="" Then
			sReefer=rs("BRFFlag")
		End If
		If sReefer = "R" Then
			sReefer = "���[�t�@�[:��"
		Else
			sReefer = "���[�t�@�[:�|"
		End If

		' ���o�\��
		Dim sCarryOut, sCarryOutFlg
		Do While Not rs.Eof
			sCarryOutFlg = CanCarryOut(rs)
			If sCarryOutFlg<>" " Then
				If sCarryOutFlg="Y" Then
					sCarryOut = "���o�F��"
				Else
					sCarryOut = "���o�F��"
				End If
				rs.MoveNext
			Else
				sCarryOut = "���o�F�~"
				Exit Do
			End If
		Loop
	End If
	rs.Close
End If
conn.Close

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If vFlg="1" Then
	If sErrMsg<>"" Then
		WriteLogM oFS, "Unknown", "2201", "�g��-�R���e�i�ԍ��Ɖ�", "10",sPhoneType, Ucase(vCtnoE) & "/" & vCtnoS & "," & "���͓��e�̐���:1(���)" & sErrMsg
	Else
		WriteLogM oFS, "Unknown", "2201", "�g��-�R���e�i�ԍ��Ɖ�", "10",sPhoneType, Ucase(vCtnoE) & "/" & vCtnoS & "," & "���͓��e�̐���:0(������)"
		WriteLogM oFS, "Unknown", "2202", "�g��-�R���e�i�ڍ�", "00",sPhoneType, CntNo & ","
	End If
Else
	WriteLogM oFS, "Unknown", "2205", "�g��-�R���e�i�ڍ�(BL)", "00",sPhoneType, CntNo & ","
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="�R���e�i�ԍ��Ɖ�">
	<center>
	�y���Ŕԍ��Ɖ�z<br><br>
<%
	If sErrMsg <> "" Then
		If sErrOpt <> "" Then
%>
			<center>
			<%=sErrOpt%><br>
<%
		End If
%>
		<center>
		<%=sErrMsg%><br>
<%
	Else
%>
		<center>
		<%=CntNo%><br>
		<center>
		<%=sCarryOut%><br>
		<center>
		<%=sPlace%><br>
		<center>
		---(�ȉ��ڍ�)---<br>
		<%=sHeight%><br>
		<%=sReefer%><br>
		<%=sDanger%><br>
		<%=sArriveTime%><br>
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
		<%=GetTitleTag("�R���e�i�ԍ��Ɖ�")%>
	</head>
	
	<body>
	<center>
	�y���Ŕԍ��Ɖ�z
	<hr>
<%
	If sErrMsg <> "" Then
		If sErrOpt <> "" Then
%>
			<%=sErrOpt%><br>
<%
		End If
%>
		<%=sErrMsg%><br><br>
<%
	Else
%>
		<%=CntNo%><br>
		<%=sCarryOut%><br>
		<%=sPlace%><br>
		---(�ȉ��ڍ�)---<br>
<%
		If sPhoneType <> "P" Then
			'PC�ȊO�͍��l(PC�͉�ʂ��L������̂ō��l�߂��Ȃ�)
%>
			</center>
<%
		End If
%>
		<%=sHeight%><br>
		<%=sReefer%><br>
		<%=sDanger%><br>
		<%=sArriveTime%><br>
		<center>
<%
	End If
%>
	<form action="index.asp" method="get">
		<input type="submit" value="�ƭ�">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>