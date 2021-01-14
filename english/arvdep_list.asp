<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<%
	'表示行数
	dispcount = 400

	'表示行数を超えた場合のエラーメッセージ
	errormessage1 = "表示件数は最大" & dispcount & "件までです。"
	errormessage2 = "検索条件に対する着離岸情報が存在しません。"
	errormessage3 = "入港日の開始日が不正です。"
	errormessage4 = "入港日の終了日が不正です。"

	com = Trim(Request.Form("com"))
	route = Trim(Request.Form("route"))
	voyage = Trim(Request.Form("voyage"))
	fromdate = Trim(Request.Form("from"))
	todate = Trim(Request.Form("to"))
	vessel = Trim(Request.Form("vessel"))
%>


<html>
<head>
	<title>着離岸情報照会</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<style type="text/css">
	<!--
		/* 着岸実績・離岸実績 */
		font.result{
			color:#ff0000;
			font-size:12px;
		}

		/* 着岸実績・離岸実績 */
		td.result{
			font-size:12px;
		}
	-->
	</style>
</head>

<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- <center> -->

<table border="0">
  <tr>
  <td width="110">
  </td>
  <td>

<table border="1" cellpadding="2" cellspacing="0">

<!-- 項目 -->
<!-- <form name="relistdata"> -->
<!-- 一覧 -->
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
	sql = sql & " WHERE DocNum IS NOT NULL AND (ShipPortSch='香椎４岸' or ShipPortSch='香椎５岸' or ShipPortSch='ＩＣ６岸') "

	If com<>"" Then
		sql = sql & " AND LineCode='" & ToSQLFormat(com) & "'"
	End If
	If route<>"" Then
		sql = sql & " AND RouteNm='" & ToSQLFormat(route) & "'"
	End If
	If voyage<>"" Then
		sql = sql & " AND Voyage='" & ToSQLFormat(voyage) & "'"
	End If
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

			' File System Object の生成
			Set fs=Server.CreateObject("Scripting.FileSystemObject")

			' 着離岸情報照会（アクセス件数用）船名
			WriteLog fs, "d101","着離岸情報照会","03", ","

		Case "2"
			sql = sql & " ORDER BY ReachSch"

			' File System Object の生成
			Set fs=Server.CreateObject("Scripting.FileSystemObject")

			' 着離岸情報照会（アクセス件数用）船名
			WriteLog fs, "d101","着離岸情報照会","04", ","

		Case "3"
			sql = sql & " ORDER BY ShipPortSch Desc, ReachSch"

			' File System Object の生成
			Set fs=Server.CreateObject("Scripting.FileSystemObject")

			' 着離岸情報照会（アクセス件数用）検索実行
			WriteLog fs, "d101","着離岸情報照会","02", ","

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

		Response.Write "<tr bgcolor='#FFFFE0'>" & CRLF

		Response.Write "<td width='110' class='result' align=' left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("CallSign")), 9) & "<br>" & FormatOptionalDigit(Trim(rsd("VslName")), 15) & "</td>" & CRLF
		Response.Write "<td width='30'class='result' align=' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("Voyage")), 5) & "<br></td>" & CRLF
		Response.Write "<td width='50'class='result' align=' align='left' valign='top' nowrap>" & ToNumberFormat(rsd("ShGweight")) & "<br>" & ToNumberFormat(rsd("ShLength")) & "</td>" & CRLF

		'抜港
		If rsd("BakkouFlg") ="1" Then
			Response.Write "<td width='55'class='result' ><br></td>" & CRLF
			Response.Write "<td width='70'class='result' ><br></td>" & CRLF
			Response.Write "<td width='70'class='result' ><br></td>" & CRLF
		Else
			Response.Write "<td width='55'class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("ShipPortSch")), 8) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipPortRes")), 8) & "</td>" & CRLF

			'着岸予定、着岸実績（着岸日時（決定）がセットされている場合は、着岸日時（予定）に着岸日時（決定）を表示する）
			If rsd("ReachDec") <>"" Then
				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachDec"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
			Else
				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("ReachSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("ReachRes"), 11) & "</font></td>" & CRLF
			End If

'20050715 Mod Start
'			Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF

			'離岸予定、離岸実績（離岸日時（決定）がセットされている場合は、離岸日時（予定）に離岸日時（決定）を表示する）
			If rsd("LeaveDec") <>"" Then
				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveDec"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
			Else
				Response.Write "<td width='70'class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("LeaveSch"), 11) & "<br><font class='result'>" & DispDateTime(rsd("LeaveRes"), 11) & "</font></td>" & CRLF
			End If
' 20050715 Mod End
		End If

		Response.Write "<td width='50'  class='result' align='left' valign='top' nowrap>" & InsertReturnCodeAtEveryOptionalDigit(FormatOptionalDigit(Trim(rsd("RouteNm")), 16), 8) & "<br></td>" & CRLF
		Response.Write "<td width='60'  class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LastRouteNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("NextRouteNm")), 10) & "</td>" & CRLF
		Response.Write "<td width='120' class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("LineCdNm")), 16) & "<br>" & FormatOptionalDigit(Trim(rsd("ShipAgency")), 16) & "</td>" & CRLF
		Response.Write "<td width='80'  class='result' align='left' valign='top' nowrap>" & FormatOptionalDigit(Trim(rsd("WkCompanyNm")), 10) & "<br>" & FormatOptionalDigit(Trim(rsd("TelNum")), 11) & "</td>" & CRLF
		Response.Write "<td width='70'  class='result' align='left' valign='top' nowrap>" & DispDateTime(rsd("RecvDate"), 11) & "<br></td>" & CRLF

		Response.Write "</tr>" & CRLF

		rsd.MoveNext
	Loop

	If listcount < 1 Then
		Response.Write "<script language='JavaScript'>window.alert('" & errormessage2 & "');</script>" & CRLF
	End If

	rsd.Close
	conn.Close
%>
<!-- /一覧 -->
<!-- </form> -->

</table>

  </td>
  </tr>
</table>

<!-- </center> -->
</body>
</html>


