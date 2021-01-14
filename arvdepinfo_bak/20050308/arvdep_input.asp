<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<html>

<head>
	<title>着離岸情報照会</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<SCRIPT src="calendarlay.js" Language="JavaScript" ></SCRIPT>

	<SCRIPT Language="JavaScript" >

	//--初期表示の為、現在日付から検索を行う。
	function getRelData() {

		document.relistf.submit();
	}

	//--現在日時を表示する(2005/2/11)
	function DispRelData() {

		mflg = false;

		from = document.relistf.from.value;
		to = document.relistf.to.value;

		// 月単位の条件の場合
		if ( from.length <= 7 )	{
			mflg = true;

			if ( from.length <= 6 )	{
				ctmp = from.substr(0,5) + "0" + from.substr(5,1);
			}
			else {
				ctmp = from;
			}

			mtmp1 = ctmp + "/01";
		}

		if ( to.length <= 7 )	{
			mflg = true;

			if ( to.length <= 6 )	{
				ctmp = to.substr(0,5) + "0" + to.substr(5,1);
			}
			else {
				ctmp = to;
			}

			mtmp2 = ctmp + "/01";
		}

		if ( mflg == true ) {
			if ( chkDate( mtmp1 ) == false ) {
				window.alert("入港日の開始日が不正です。");
				document.relistf.from.value = from;
				document.relistf.to.value = to;
				return;
			}

			if ( chkDate( mtmp2 ) == false ) {
				window.alert("入港日の終了日が不正です。");
				document.relistf.from.value = from;
				document.relistf.to.value = to;
				return;
			}
		}
		else {

			if ( document.relistf.from.value !="" ) { 
				if ( chkDate( document.relistf.from.value) == false ) {
					window.alert("入港日の開始日が不正です。");
					document.relistf.from.value = from;
					document.relistf.to.value = to;
					return;
				}
			}

			if ( document.relistf.to.value !="" ) { 
				if ( chkDate( document.relistf.to.value) == false ) {
					window.alert("入港日の終了日が不正です。");
					document.relistf.from.value = from;
					document.relistf.to.value = to;
					return;
				}
			}
		}

		document.relistf.submit();

	}

	function chkDate (yyyymmdd) {

		//-- format yyyy/mm/dd
		midx = 7;
		var years = yyyymmdd.substr(0,4);
		var months = yyyymmdd.substr(5,2);
		if ( months.substr(0,1) == "0" ) {
			midx++;
			months = months.substr(1,1);
		}
		else {
			if (months.substr(1,1) == "/") {
				months = months.substr(0,1);
			}
			else {
				midx++;
				months = months.substr(0,2);
			}
		}

		var days = yyyymmdd.substr(midx,2);
		if (days.substr(0,1) == "0") {
			days = days.substr(1,1);
		}
		else {
			if (days.substr(1,1) == "/") {
				days = months.substr(0,1);
			}
			else {
				days = days.substr(0,2);
			}
		}

		var flag = true;
		years = parseInt(years);
		months = parseInt(months) - 1;
		days = parseInt(days);

		var dates = new Date(years,months,days);
		if (dates.getYear() < 1900) {
			if (years != dates.getYear() + 1900) {
				flag = false;
			}
		}
		else {
			if (years != dates.getYear()) {
				flag = false;
			}
		}

		if (months != dates.getMonth()) {
			flag = false;
		}

		// 日付チェック有り
		if (days != dates.getDate()) {
			flag = false;
		}

		return flag;
	}

	</SCRIPT>

    <style type="text/css">
	<!--
		/* 太字リンク */
		font.menulink{
			color:#0055ff;
			font-size:14px;
			font-weight:bold;
		}

		/* 検索項目 */
		td.search{
			width:    100px;
			height:    23px;
			font-size: 14px;
			color:#ffffff;
			font-weight:bold;
			background-color:#000099;
			padding: 3px 5px 3px 5px;
		}
		/* 次航、入港日 */
		input.search{
			width:    150px;
			height:    23px;
			font-size: 14px;
		}
		/* 運航船社、航路 */
		select.search{
			width:    150px;
			height:    23px;
			font-size: 14px;
		}

		/* 説明文 */
		td.explain{
			font-size:12px;
			color:#000000;
			font-weight:bold;
		}
	-->
	</style>
</head>

<body onload="getRelData()" bgcolor="#dee1ff" text="#000000" link="#3300ff" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><!-- ヘッダ -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form name="relistf" method="post" action="./arvdep_list.asp" target="list">
	<tr>
	  <td rowspan="2">
		<IMG height=73 src="gif/arvdep.gif" width=507>
	  </td>
	  <td height="19" bgcolor="#000099" align="right">
		<IMG height=19 src="gif/logo_hits_ver2_1.gif" width=300>
	  </td>
	</tr>
	<tr>
	  <td align="right" width="100%" height="48">
		<table border="0" cellpadding="0" cellspacing="0">
		  <tr>
			<td nowrap>
				<a href="http://www.port-of-hakata.or.jp/business/cargo/index_list.htm" tabindex="1" target="_top"><font class="menulink">航路一覧</font></a>
			</td>
			<td width="50"></td>
		  </tr>
		</table>
	  </td>
	</tr>
  </table>

<center>
<BR><!-- /ヘッダ -->

  <table border="0" cellpadding="3" cellspacing="0"><!-- 検索条件 -->
	<tr>
	  <td  align="left" valigh="top">
		  <table border="0" cellpadding="0" cellspacing="0">
			<tr>
			  <td align="left" valign="top" nowrap>
				  <table border="0" cellpadding="0" cellspacing="0">
					<tr>
					  <td>
						<table border="1" cellpadding="0" cellspacing="0" frame="hsides" bordercolor="#dee1ff" bordercolordark="#dee1ff" bordercolorlight="#dee1ff">
						  <td class="search" align="left" valign="center" nowrap>
							運航船社
						  </td>
						</table>
					  </td>
					  <td align="left" valign="center">
						<select class="search" name="com" tabindex="2" style="WIDTH: 303px">
							<option value=""/>
<%
	ConnectSvr conn, rsd

	sql = "select LineCd,LineCdNm from sLineCmpany order by LineCdNm"
	rsd.Open sql, conn, 0, 1, 1

	Do While Not rsd.EOF
		Response.Write "<option value='" & Trim(rsd("LineCd")) & "'>" & Trim(rsd("LineCdNm")) & "</option>"
		rsd.MoveNext
	Loop

    rsd.Close
%>
						</select>
					  </td>
					  <td>
&nbsp;&nbsp;
					  </td>
					  <td>
						<table border="1" cellpadding="0" cellspacing="0" frame="hsides" bordercolor="#dee1ff" bordercolordark="#dee1ff" bordercolorlight="#dee1ff">
						  <td class="search" align="left" valign="center" nowrap>
							次航
						  </td>
						</table>
					  </td>
					  <td align="left" valign="center">
						<input class="search" name="next" maxlength="5" tabindex="3" style="WIDTH: 50px; HEIGHT: 23px" size=5>
					  </td>
					</tr>
				  </table>
			  </td>
			</tr>
			<tr><td height="5"></td>
			<tr>
			  <td align="left" valign="top" nowrap>
				  <table border="0" cellpadding="0" cellspacing="0">
					<tr>
					  <td>
						<table border="1" cellpadding="0" cellspacing="0" frame="hsides" bordercolor="#dee1ff" bordercolordark="#dee1ff" bordercolorlight="#dee1ff">
						  <td class="search" align="left" valign="center" nowrap>
							航路
						  </td>
						</table>
					  </td>
					  <td align="left" valign="center">
						<select class="search" name="voyage" tabindex="4" style="WIDTH: 180px">
							<option value=""/>
<%
	sql = "select RouteNm,RouteAlp from sShipRoute order by RouteAlp"
	rsd.Open sql, conn, 0, 1, 1

	Do While Not rsd.EOF
		tmpname = Trim(rsd("RouteNm"))
		Response.Write "<option value='" & tmpname & "'>" & tmpname & "</option>"
		rsd.MoveNext
	Loop

    rsd.Close

	conn.Close
%>
						</select>
					  </td>
					  <td>
&nbsp;&nbsp;
					  </td>
					  <td>
						<table border="1" cellpadding="0" cellspacing="0" frame="hsides" bordercolor="#dee1ff" bordercolordark="#dee1ff" bordercolorlight="#dee1ff">
						  <td class="search" align="left" valign="center" nowrap>
							入港日
						  </td>
						</table>
					  </td>
					  <td align="left" valign="top" nowrap>
						  <table border="0" cellpadding="0" cellspacing="0">
							<tr>
							  <td align="left" valign="center" nowrap>
<%
	Response.Write "								<input class='search' type='text' name='from' value='" & ChgYMDStr4(GetYMDStr(Now)) & "' maxlength='10' tabindex='5' style='WIDTH: 85px; HEIGHT: 23px' size=10>" & CRLF
%>
								<input type="button" value="*" onclick="wrtCalendarLay(this.form.from,event)">
								<b>～</b>
<%
	Response.Write "								<input class='search' type='text' name='to' value='" & ChgYMDStr4(GetYMDStr(Now)) & "' maxlength='10' tabindex='6' style='WIDTH: 85px; HEIGHT: 23px' size=10>" & CRLF
%>
								<input type="button" value="*" onclick="wrtCalendarLay(this.form.to,event)">
							  </td>
							</tr>
						  </table>
					  </td>
					</tr>
				  </table>
			  </td>
			</tr>
		  </table>
	  </td>
	</tr><!-- /検索条件 --><!-- 説明文 -->
	<tr>
	  <td align="left" valign="top">
		  <table border="0" cellpadding="2" cellspacing="0">
			<tr>
			  <td class="explain" colspan="2" align="left" nowrap>
				※入港日について
			  </td>
			</tr>
			<tr>
			  <td width="15" rowspan="3">
			  <td class="explain" align="left" nowrap>
				・着岸予定が該当するものをすべて表示します。
			  </td>
			</tr>
			<tr>
			  <td class="explain" align="left" nowrap>
				・月、または、日までの指定が可能です(2005/9)(2005/9/25)。
			  </td>
			</tr>
		  </table>
	  </td>
	  <td align="right" valign="top">
		<A tabIndex=7 href="javascript:DispRelData()"><font class="menulink">検索実行</font></A> 
	  </td>
	</tr><!-- /説明文 --></FORM>
</table></center>

</body>
</html>
