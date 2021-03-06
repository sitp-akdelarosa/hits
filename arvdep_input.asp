<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<html>

<head>
	<title>着離岸情報照会</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!-- add by nics 2015.03.11 -->
	<link href="./reset.css" rel="stylesheet" type="text/css">
<!-- end of add by nics 2015.03.11 -->
	<SCRIPT src="calendarlay.js" Language="JavaScript" ></SCRIPT>

	<SCRIPT Language="JavaScript" >

	//--初期表示の為、現在日付から検索を行う。
	function getRelData() {

		//--検索実行ボタンを処理する（DispRelData()）
		document.relistf.submit();

	}

	//--検索実行ボタンを処理する
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

		//-- document.relistf.vessel.value="0";
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
		/* 運航行会社、航路 */
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

		/* ソート用リンク */
		a.sortlink{
			color:#0055ff;
			font-size:12px;
		}

		/* 一覧 */
		td.listtitle{
			font-size: 12px;
			color:#000000;
			background-color:#aaaaff;
		}

/* add by nics 2015.03.11 */
		table#title {
			border-style: solid;
			border-color: #ffffff;
			border-width: 1px 0px 0px 1px;
		}

		table#title td {
			padding: 3px 3px 3px 3px;
			border-style: solid;
			border-color: #ffffff;
			border-width: 0px 1px 1px 0px;
		}
/* end of add by nics 2015.03.11 */
	-->
	</style>
</head>

<body bgcolor="#dee1ff" text="#000000" link="#3300ff" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><!-- ヘッダ -->

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form name="relistf" method="post" action="./arvdep_list.asp" target="list">

	<input type="hidden" name="vessel" value="">

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
				<a href="http://www.port-of-hakata.or.jp/business/cargo/index.html" tabindex="1" target="_blank"><font class="menulink">航路一覧</font></a>
			</td>
			<td width="50"></td>
		  </tr>
		</table>
	  </td>
	</tr>
  </table>
<!-- <center> -->
<BR><!-- /ヘッダ -->

<table border="0">
 <tr>
 <td width="110">
 </td>
 <td>

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
						<input class="search" name="voyage" maxlength="5" tabindex="3" style="WIDTH: 50px; HEIGHT: 23px" size=5>
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
						<select class="search" name="route" tabindex="4" style="WIDTH: 180px">
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
								<b>〜</b>
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
				・着岸予定、または、離岸予定が該当するものをすべて表示します。
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
<!-- mod by nics 2015.03.11
		<A tabIndex=7 href="javascript:document.relistf.vessel.value='3';javascript:DispRelData()"><font class="menulink">検索実行</font></A>  -->
		<A tabIndex=7 href="javascript:document.relistf.vessel.value='3';javascript:DispRelData()"  style="white-space: nowrap;"><font class="menulink">検索実行</font></A> 
<!-- end of mod by nics 2015.03.11 -->
	  </td>
	</tr><!-- /説明文 --></FORM>
</table>

</td>
</tr>
</table>

<!-- </center> -->

<br>
<!-- <center> -->

<table border="0">
  <tr>
  <td width="110">
  </td>
  <td>

<!-- 受信時間表示 -->
<%
	' SQL
	sql = "SELECT max(RecvDate) as RecvDate FROM sReLeaveDat;"

	'Response.Write sql

	ConnectSvr conn, rsd
	rsd.Open sql, conn, 0, 1, 1

	Response.Write "<td align='left' valign='top' nowrap>" & DispDateTime2(rsd("RecvDate"), 20) & "　現在の情報（約30分間隔で更新中）<br><font class='result'>" & "</font></td>" & CRLF

	rsd.Close
	conn.Close
%>
<!-- 受信時間表示 -->
 </td>
 </tr>
</table>

<!-- mod by nics 2015.03.11
<table border="0">
  <tr>
  <td width="110"></td>
  <td>

<table border="1" cellpadding="2" cellspacing="0">
	<tr>
	  <td width="110" class="listtitle" align="left" valign="top" nowrap>
		コールサイン<BR><a class="sortlink" href="javascript:document.relistf.vessel.value='1';document.relistf.submit();">船　名</a><BR>　
	  </td>
	  <td width="30" class="listtitle" align="left" valign="top" nowrap>
		次航
	  </td>
	  <td width="50" class="listtitle" align="left" valign="top" nowrap>
		総トン数<BR>全　長<BR>　
	  </td>
	  <td width="55" class="listtitle" align="left" valign="top" nowrap>
		予定船席<BR>実績船席<BR>　
	  </td>
	  <td width="70" class="listtitle" align="left" valign="top" nowrap>
		<a class="sortlink" href="javascript:document.relistf.vessel.value='2';document.relistf.submit();">着岸予定</a><BR>着岸実績
	  </td>
	  <td width="70" class="listtitle" align="left" valign="top" nowrap>
		離岸予定<BR>離岸実績<BR>　
	  </td>
	  <td width="50" class="listtitle" align="left" valign="top" nowrap>
		航路<BR>
	  </td>
	  <td width="60" class="listtitle" align="left" valign="top" nowrap>
		前　港<BR>次　港<BR>　
	  </td>
	  <td width="120" class="listtitle" align="left" valign="top" nowrap>
		運航船社<BR>船舶代理店<BR>　
	  </td>
	  <td width="80" class="listtitle" align="left" valign="top" nowrap>
		オペレータ<BR>連絡先<BR>　
	  </td>
	  <td width="70" class="listtitle" align="left" valign="top" nowrap>
		更新時間<BR>　
	  </td>
	</tr>
  </table>

  </td>
  </tr>
</table> -->
<table border="0" style="width: 863px;margin-left: 115px;" class="reset">
  <tr class="reset">
  <td width="110" class="reset"></td>
  <td  class="reset">

<table border="1" cellpadding="2" cellspacing="0" class="reset" id="title" style="width: 863px">
<!-- 項目 -->
	<tr class="reset">
	  <td width="110" class="reset listtitle" align="left" valign="top" nowrap>
		コールサイン<BR><a class="sortlink" href="javascript:document.relistf.vessel.value='1';document.relistf.submit();">船　名</a><BR>　
	  </td>
	  <td width="30" class="reset listtitle" align="left" valign="top" nowrap>
		次航
	  </td>
	  <td width="60" class="reset listtitle" align="left" valign="top" nowrap>
		総トン数<BR>全　長<BR>　
	  </td>
	  <td width="55" class="reset listtitle" align="left" valign="top" nowrap>
		予定船席<BR>実績船席<BR>　
	  </td>
	  <td width="70" class="reset listtitle" align="left" valign="top" nowrap>
		<a class="sortlink" href="javascript:document.relistf.vessel.value='2';document.relistf.submit();">着岸予定</a><BR>着岸実績
	  </td>
	  <td width="70" class="reset listtitle" align="left" valign="top" nowrap>
		離岸予定<BR>離岸実績<BR>　
	  </td>
	  <td width="50" class="reset listtitle" align="left" valign="top" nowrap>
		航路<BR>
	  </td>
	  <td width="70" class="reset listtitle" align="left" valign="top" nowrap>
		前　港<BR>次　港<BR>　
	  </td>
	  <td width="120" class="reset listtitle" align="left" valign="top" nowrap>
		運航船社<BR>船舶代理店<BR>　
	  </td>
	  <td width="80" class="reset listtitle" align="left" valign="top" nowrap>
		オペレータ<BR>連絡先<BR>　
	  </td>
	  <td width="70" class="reset listtitle" align="left" valign="top" nowrap>
		更新時間<BR>　
	  </td>
	</tr>
<!-- /項目 -->
  </table>
<!-- </center> -->

  </td>
  </tr>
</table>
<!-- end of mod by nics 2015.03.11
<!-- add by nics 2015.03.11 -->
<iframe src="./arvdep_list.asp" name="list" noresize scrolling="no"  frameborder="0" width="1000px"  height="441px" class="reset" marginwidth="0" marginheight="0" style="margin-top:5px;"></iframe>
<!-- end of add by nics 2015.03.11 -->
</body>
</html>


