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
	nextroute = Trim(Request.Form("next"))
	voyage = Trim(Request.Form("voyage"))
	fromdate = Trim(Request.Form("from"))
	todate = Trim(Request.Form("to"))
	vessel = Trim(Request.Form("vessel"))
%>


<html>
<head>
	<title>�����ݏ��Ɖ�</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<style type="text/css">
	<!--
		/* �ꗗ */
		td.listtitle{
			font-size: 12px;
			color:#000000;
			background-color:#aaaaff;
			padding: 3px 5px 3px 5px;
		}

		/* �\�[�g�p�����N */
		a.sortlink{
			color:#0055ff;
			font-size:12px;
		}

		/* ���ݎ��сE���ݎ��� */
		font.result{
			color:#ff0000;
			font-size:12px;
		}

		/* ���ݎ��сE���ݎ��� */
		td.result{
			font-size:12px;
		}
	-->
	</style>
</head>

<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>

<!-- ��M���ԕ\�� -->
<%
	' SQL
	sql = "SELECT max(RecvDate) as RecvDate FROM sReLeaveDat;"

'Response.Write sql

	ConnectSvr conn, rsd
	rsd.Open sql, conn, 0, 1, 1

	Response.Write "<td align='left' valign='top' nowrap>" & DispDateTime2(rsd("RecvDate"), 20) & "�@���݂̏��i��30���Ԋu�ōX�V���j<br><font class='result'>" & "</font></td>" & CRLF

    rsd.Close
	conn.Close
%>
<!-- ��M���ԕ\�� -->

<table border="1" cellpadding="2" cellspacing="0">

<!-- ���� -->
  <form method="post" action="./arvdep_list.asp">
	<input type="hidden" name="vessel" value="">
	<tr>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�R�[���T�C��<BR><a class="sortlink" href="javascript:document.forms[0].vessel.value='1';document.forms[0].submit();">�D�@��</a>
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		���q
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		���g����<BR>�S�@��
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�\��D��<BR>���ёD��
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		<a class="sortlink" href="javascript:document.forms[0].vessel.value='2';document.forms[0].submit();">���ݗ\��</a><BR>���ݎ���
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		���ݗ\��<BR>���ݎ���
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�q�H
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�O�@�`<BR>���@�`
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�^�q�D��<BR>�D���㗝�X
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�I�y���[�^<BR>�A����
	  </td>
	  <td class="listtitle" align="left" valign="top" nowrap>
		�X�V����
	  </td>
	</tr>
<!-- /���� -->

<!-- �ꗗ -->
<%
	Response.Write "<input type='hidden' name='com' value='" & Trim(Request.Form("com")) & "'>" & CRLF
	Response.Write "<input type='hidden' name='next' value='" & Trim(Request.Form("next")) & "'>" & CRLF
	Response.Write "<input type='hidden' name='voyage' value='" & Trim(Request.Form("voyage")) & "'>" & CRLF
	Response.Write "<input type='hidden' name='from' value='" & Trim(Request.Form("from")) & "'>" & CRLF
	Response.Write "<input type='hidden' name='to' value='" & Trim(Request.Form("to")) & "'>" & CRLF

	' SQL
	sql = "SELECT r.*, l.LineCdNm, w.WkCompanyNm, w.TelNum"
	sql = sql & " FROM sReLeaveDat r"
	sql = sql & " LEFT OUTER JOIN sLineCmpany l ON (l.LineCd=r.LineCode)"
	sql = sql & " LEFT OUTER JOIN sWkCompany w  ON (w.WkCompanyCd=r.WkCompanyCd)"
	sql = sql & " WHERE DocNum IS NOT NULL AND (ShipPortSch<>'�����S��' AND ShipPortSch<>'�h�b�T��') "

	If com<>"" Then
		sql = sql & " AND LineCode='" & ToSQLFormat(com) & "'"
	End If
	If nextroute<>"" Then
		sql = sql & " AND NextRouteNm='" & ToSQLFormat(nextroute) & "'"
	End If
	If voyage<>"" Then
		sql = sql & " AND RouteNm='" & ToSQLFormat(voyage) & "'"
	End If
	If fromdate<>"" and todate<>"" Then
		fromary = Split(fromdate, "/")
		toary = Split(todate, "/")

		If IsRightDateFormat(fromary)=True Then
			Select Case UBound(fromary)
				Case 1
					sql = sql & " AND ReachSch>='" & fromdate & "/1 00:00:00.000'"
				Case 2
					sql = sql & " AND ReachSch>='" & fromdate & " 00:00:00.000'"
				Case Else
			End Select
		ELSE
			Response.Write "<script language='JavaScript'>window.alert('" & errormessage3 & "');</script>" & CRLF
		End If

		If IsRightDateFormat(toary)=True Then
			Select Case UBound(toary)
				Case 1
					sql = sql & " AND ReachSch<'" & FirstDayOfNextMonth(toary(0), toary(1)) & "'"
				Case 2
					sql = sql & " AND ReachSch<='" & todate & " 23:59:59.000'"
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
		Case "2"
			sql = sql & " ORDER BY ReachSch"
		Case Else
			sql = sql & " ORDER BY ReachSch DESC"
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

		Response.Write "<tr bgcolor='#FFFFE0'>" & CRLF

		Response.Write "<td class='result' align=' left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("CallSign")), 9) & "<br>" & FormatOptionalDigit(Trim(rsd("VslName")), 15) & "</td>" & CRLF
		Response.Write "<td class='result' align=' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("Voyage")), 5) & "<br></td>" & CRLF
		Response.Write "<td class='result' align=' align='left' valign='top' nowrap>" & ToNumberFormat(rsd("ShGweight")) & "<br>" & ToNumberFormat(rsd("ShLength")) & "</td>" & CRLF

		'���`
		If rsd("BakkouFlg") ="1" Then
			Response.Write "<td class='result' ><br></td>" & CRLF
			Response.Write "<td class='result' ><br></td>" & CRLF
			Response.Write "<td class='result' ><br></td>" & CRLF
		Else
			Response.Write "<td class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("ShipPortSch")), 8) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipPortRes")), 8) & "</td>" & CRLF
			Response.Write "<td class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
			Response.Write "<td class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
		End If

		Response.Write "<td class='result' align='left' valign='top' nowrap>" & InsertReturnCodeAtEveryOptionalDigit(FormatOptionalDigit(Trim(rsd("RouteNm")), 16), 8) & "<br></td>" & CRLF
		Response.Write "<td class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LastRouteNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("NextRouteNm")), 10) & "</td>" & CRLF
		Response.Write "<td class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LineCdNm")), 16) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipAgency")), 16) & "</td>" & CRLF
		Response.Write "<td class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("WkCompanyNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("TelNum")), 11) & "</td>" & CRLF
		Response.Write "<td class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("RecvDate"), 11) & "<br></td>" & CRLF

		Response.Write "</tr>" & CRLF

		rsd.MoveNext
	Loop

	If listcount < 1 Then
		Response.Write "<script language='JavaScript'>window.alert('" & errormessage2 & "');</script>" & CRLF
	End If

    rsd.Close
	conn.Close
%>
<!-- /�ꗗ -->
  </form>

</table>

</center>
</body>
</html>
