<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<%
	'�\���s��
	dispcount = 400

	'�\���s���𒴂����ꍇ�̃G���[���b�Z�[�W
	errormessage1 = "�\�������͍ő�" & dispcount & "���܂łł��B"
	errormessage2 = "���������ɑ΂��钅���ݏ�񂪑��݂��܂���B"
	errormessage3 = "���`���̊J�n�����s���ł��B"
	errormessage4 = "���`���̏I�������s���ł��B"

	com = Trim(Request.Form("com"))
	route = Trim(Request.Form("route"))
	voyage = Trim(Request.Form("voyage"))
	fromdate = Trim(Request.Form("from"))
	todate = Trim(Request.Form("to"))
	vessel = Trim(Request.Form("vessel"))
%>


<html>
<head>
	<title>�����ݏ��Ɖ�</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!-- add by nics 2015.03.11 -->
	<link href="./reset.css" rel="stylesheet" type="text/css">
<!-- end of add by nics 2015.03.11 -->
	<style type="text/css">
	<!--
		/* ���ݎ��сE���ݎ��� */
		font.result{
			color:#ff0000;
			font-size:12px;
		}

		/* ���ݎ��сE���ݎ��� */
		td.result{
			font-size:12px;
		}

/* add by nics 2015.03.11 */
		table#list {
			border-style: solid;
			border-color: #aca899;
			border-width: 1px 0px 0px 1px;
		}

		table#list td {
			padding: 3px 3px 2px 3px;
			border-style: solid;
			border-color: #aca899;
			border-width: 0px 1px 1px 0px;
		}
/* end of add by nics 2015.03.11 */
	-->
	</style>
</head>

<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- <center> -->

<!-- mod by nics 2015.03.11
<table border="0">
  <tr>
  <td width="110">
  </td>
  <td>

<table border="1" cellpadding="2" cellspacing="0"> -->
<div style="height: 441px;overflow-x :hidden; overflow-y: auto; width: 882px;margin-left: 115px;" class="reset">
<table border="0" style="width: 863px;" class="reset">
  <tr class="reset">
  <td width="110" class="reset"></td>
  <td class="reset">
<table class="reset" id="list" style="width:863px;">
<!-- end of mod by nics 2015.03.11 -->

<!-- ���� -->
<!-- <form name="relistdata"> -->
<!-- �ꗗ -->
<%
Response.Write "<input type='hidden' name='com' value='" & Trim(Request.Form("com")) & "'>" & CRLF
Response.Write "<input type='hidden' name='route' value='" & Trim(Request.Form("route")) & "'>" & CRLF
Response.Write "<input type='hidden' name='voyage' value='" & Trim(Request.Form("voyage")) & "'>" & CRLF
Response.Write "<input type='hidden' name='from' value='" & Trim(Request.Form("from")) & "'>" & CRLF
Response.Write "<input type='hidden' name='to' value='" & Trim(Request.Form("to")) & "'>" & CRLF
Response.Write "<input type='hidden' name='vessel' value='" & Trim(Request.Form("vessel")) & "'>" & CRLF

	' SQL
	sql = "SELECT r.*, l.LineCdNm, w.WkCompanyNm, w.TelNum"
	sql = sql & " FROM sReLeaveDat r"
	sql = sql & " LEFT OUTER JOIN sLineCmpany l ON (l.LineCd=r.LineCode)"
	sql = sql & " LEFT OUTER JOIN sWkCompany w  ON (w.WkCompanyCd=r.WkCompanyCd)"
	sql = sql & " WHERE DocNum IS NOT NULL AND (ShipPortSch='���łS��' or ShipPortSch='���łT��' or ShipPortSch='�h�b�U��' or ShipPortSch='�����T��' or ShipPortSch='�h�b�T��' or ShipPortSch='�h�b�V��') "

	If com<>"" Then
		sql = sql & " AND LineCode='" & ToSQLFormat(com) & "'"
	End If
	If route<>"" Then
		sql = sql & " AND RouteNm='" & ToSQLFormat(route) & "'"
	End If
	If voyage<>"" Then
		sql = sql & " AND Voyage='" & ToSQLFormat(voyage) & "'"
	End If

	' �����\���̏ꍇ�́A���ݓ��t���Z�b�g����B
	If fromdate="" and todate="" Then

		fromdate = ChgYMDStr4(GetYMDStr(Now))
		todate   = ChgYMDStr4(GetYMDStr(Now))

	End If

'		Response.Write "<script language='JavaScript'>window.alert(" & fromdate & ");</script>" & CRLF
'		Response.Write "<script language='JavaScript'>window.alert(" & todate & ");</script>" & CRLF

	If fromdate<>"" and todate<>"" Then
		fromary = Split(fromdate, "/")
		toary = Split(todate, "/")

		If IsRightDateFormat(fromary)=True Then
			Select Case UBound(fromary)
				Case 1
					sql = sql & " AND ((ReachSch>='" & fromdate & "/1 00:00:00.000') OR (LeaveSch>='" & fromdate & "/1 00:00:00.000')) "
				Case 2
					sql = sql & " AND ((ReachSch>='" & fromdate & " 00:00:00.000') OR (LeaveSch>='" & fromdate & " 00:00:00.000')) "
				Case Else
			End Select
		ELSE
			Response.Write "<script language='JavaScript'>window.alert('" & errormessage3 & "');</script>" & CRLF
		End If

		If IsRightDateFormat(toary)=True Then
			Select Case UBound(toary)
				Case 1
					sql = sql & " AND ((ReachSch<'" & FirstDayOfNextMonth(toary(0), toary(1)) & "') OR (LeaveSch<'" & FirstDayOfNextMonth(toary(0), toary(1)) & "'))"
				Case 2
					sql = sql & " AND ((ReachSch<='" & todate & " 23:59:59.000') OR (LeaveSch<='" & todate & " 23:59:59.000'))"
				Case Else
			End Select
		ELSE
			Response.Write "<script language='JavaScript'>window.alert('" & errormessage4 & "');</script>" & CRLF
		End If
	End If
	If fromdate<>"" and todate="" Then
		fromary = Split(fromdate, "/")

		If IsRightDateFormat(fromary)=True Then
			Select Case UBound(fromary)
				Case 1
					sql = sql & " AND ReachSch>='" & fromdate & "/1 00:00:00.000'"
					sql = sql & " AND ReachSch<'" & FirstDayOfNextMonth(fromary(0), fromary(1)) & "'"
				Case 2
					sql = sql & " AND ReachSch>='" & fromdate & " 00:00:00.000'"
					sql = sql & " AND ReachSch<='" & fromdate & " 23:59:59.000'"
				Case Else
			End Select
		ELSE
			Response.Write "<script language='JavaScript'>window.alert('" & errormessage3 & "');</script>" & CRLF
		End If
	End If

	Select Case vessel
		Case "1"
			sql = sql & " ORDER BY VslName"

			' File System Object �̐���
			Set fs=Server.CreateObject("Scripting.FileSystemObject")

			' �����ݏ��Ɖ�i�A�N�Z�X�����p�j�D��
			WriteLog fs, "d101","�����ݏ��Ɖ�","03", ","

		Case "2"
			sql = sql & " ORDER BY ReachSch"

			' File System Object �̐���
			Set fs=Server.CreateObject("Scripting.FileSystemObject")

			' �����ݏ��Ɖ�i�A�N�Z�X�����p�j�D��
			WriteLog fs, "d101","�����ݏ��Ɖ�","04", ","

		Case "3"
			sql = sql & " ORDER BY ShipPortSch Desc, ReachSch"

			' File System Object �̐���
			Set fs=Server.CreateObject("Scripting.FileSystemObject")

			' �����ݏ��Ɖ�i�A�N�Z�X�����p�j�������s
			WriteLog fs, "d101","�����ݏ��Ɖ�","02", ","

		Case Else
			sql = sql & " ORDER BY ShipPortSch Desc, ReachSch"
	End Select

'Response.Write sql

	ConnectSvr conn, rsd
	rsd.Open sql, conn, 0, 1, 1

	listcount = 0
	Do While Not rsd.EOF
		listcount = listcount + 1
		If listcount > dispcount Then
			Response.Write "<script language='JavaScript'>window.alert('" & errormessage1 & "');</script>" & CRLF
			Exit Do
		End If
' mod by nics 2015.03.11
'		Response.Write "<tr bgcolor='#FFFFE0'>" & CRLF
'
'		Response.Write "<td width='110' class='result' align=' left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("CallSign")), 9) & "<br>" & FormatOptionalDigit(Trim(rsd("VslName")), 15) & "</td>" & CRLF
'		Response.Write "<td width='30'class='result' align=' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("Voyage")), 5) & "<br></td>" & CRLF
'' mod by nics 2015.03.11
''		Response.Write "<td width='50'class='result' align=' align='left' valign='top' nowrap>" & ToNumberFormat(rsd("ShGweight")) & "<br>" & ToNumberFormat(rsd("ShLength")) & "</td>" & CRLF
'		Response.Write "<td width='55'class='result' align=' align='left' valign='top' nowrap>" & ToNumberFormat(rsd("ShGweight")) & "<br>" & ToNumberFormat(rsd("ShLength")) & "</td>" & CRLF
'
'		'���`
'		If rsd("BakkouFlg") ="1" Then
'			Response.Write "<td width='55'class='result' ><br></td>" & CRLF
'			Response.Write "<td width='70'class='result' ><br></td>" & CRLF
'			Response.Write "<td width='70'class='result' ><br></td>" & CRLF
'		Else
'			Response.Write "<td width='55'class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("ShipPortSch")), 8) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipPortRes")), 8) & "</td>" & CRLF
'
'			'���ݗ\��A���ݎ��сi���ݓ����i����j���Z�b�g����Ă���ꍇ�́A���ݓ����i�\��j�ɒ��ݓ����i����j��\������j
'			If rsd("ReachDec") <>"" Then
'				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachDec"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
'			Else
'				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
'			End If
'
''20050715 Mod Start
''			Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
'
'			'���ݗ\��A���ݎ��сi���ݓ����i����j���Z�b�g����Ă���ꍇ�́A���ݓ����i�\��j�ɗ��ݓ����i����j��\������j
'			If rsd("LeaveDec") <>"" Then
'				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveDec"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
'			Else
'				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
'			End If
'' 20050715 Mod End
'		End If
'
'		Response.Write "<td width='50'  class='result' align='left' valign='top' nowrap>" & InsertReturnCodeAtEveryOptionalDigit(FormatOptionalDigit(Trim(rsd("RouteNm")), 16), 8) & "<br></td>" & CRLF
'		Response.Write "<td width='60'  class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LastRouteNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("NextRouteNm")), 10) & "</td>" & CRLF
'		Response.Write "<td width='120' class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LineCdNm")), 16) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipAgency")), 16) & "</td>" & CRLF
'		Response.Write "<td width='80'  class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("WkCompanyNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("TelNum")), 11) & "</td>" & CRLF
'		Response.Write "<td width='70'  class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("RecvDate"), 11) & "<br></td>" & CRLF
'
'		Response.Write "</tr>" & CRLF
		Response.Write "<tr bgcolor='#FFFFE0'>" & CRLF

		Response.Write "<td width='110' class='reset result' align=' left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("CallSign")), 9) & "<br>" & FormatOptionalDigit(Trim(rsd("VslName")), 15) & "</td>" & CRLF
		Response.Write "<td width='30'class='reset result' align=' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("Voyage")), 5) & "<br></td>" & CRLF
		Response.Write "<td width='60'class='reset result' align=' align='left' valign='top' nowrap>" & ToNumberFormat(rsd("ShGweight")) & "<br>" & ToNumberFormat(rsd("ShLength")) & "</td>" & CRLF

		'���`
		If rsd("BakkouFlg") ="1" Then
			Response.Write "<td width='55'class='reset result' ><br></td>" & CRLF
			Response.Write "<td width='70'class='reset result' ><br></td>" & CRLF
			Response.Write "<td width='70'class='reset result' ><br></td>" & CRLF
		Else
			Response.Write "<td width='55'class='reset result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("ShipPortSch")), 8) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipPortRes")), 8) & "</td>" & CRLF

			'���ݗ\��A���ݎ��сi���ݓ����i����j���Z�b�g����Ă���ꍇ�́A���ݓ����i�\��j�ɒ��ݓ����i����j��\������j
			If rsd("ReachDec") <>"" Then
				Response.Write "<td width='70'class='reset result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachDec"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
			Else
				Response.Write "<td width='70'class='reset result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
			End If

'20050715 Mod Start
'			Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF

			'���ݗ\��A���ݎ��сi���ݓ����i����j���Z�b�g����Ă���ꍇ�́A���ݓ����i�\��j�ɗ��ݓ����i����j��\������j
			If rsd("LeaveDec") <>"" Then
				Response.Write "<td width='70'class='reset result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveDec"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
			Else
				Response.Write "<td width='70'class='reset result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
			End If
' 20050715 Mod End
		End If

		Response.Write "<td width='50'  class='reset result' align='left' valign='top' nowrap>" & InsertReturnCodeAtEveryOptionalDigit(FormatOptionalDigit(Trim(rsd("RouteNm")), 16), 8) & "<br></td>" & CRLF
		Response.Write "<td width='70'  class='reset result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LastRouteNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("NextRouteNm")), 10) & "</td>" & CRLF
		Response.Write "<td width='120' class='reset result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LineCdNm")), 16) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipAgency")), 16) & "</td>" & CRLF
		Response.Write "<td width='80'  class='reset result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("WkCompanyNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("TelNum")), 11) & "</td>" & CRLF
		Response.Write "<td width='70'  class='reset result' align='left' valign='top' nowrap>" & DispDateTime(rsd("RecvDate"), 11) & "<br></td>" & CRLF

		Response.Write "</tr>" & CRLF
' end of mod by nics 2015.03.11

		rsd.MoveNext
	Loop

	If listcount < 1 Then
		Response.Write "<script language='JavaScript'>window.alert('" & errormessage2 & "');</script>" & CRLF
	End If

	rsd.Close
	conn.Close
%>
<!-- /�ꗗ -->
<!-- </form> -->

</table>

  </td>
  </tr>
</table>

<!-- </center> -->
</body>
</html>

