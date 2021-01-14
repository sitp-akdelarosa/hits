<%
	@LANGUAGE = VBScript
	@CODEPAGE = 932
%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst000M.asp				_/
'_/	Function	:ステータス配信一覧画面メニュー		_/
'_/	Date			:2003/12/25				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>Import Status Delivery Request
</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
//遷移
function opnewin(i){
	Fname=document.sst000MF;
	Fname.target="List";
	switch(i){
		case 1:
			Fname.action="./sst100F.asp";
			break;
		case 2:
			Win = window.open('sst200.asp', 'FConIn', 'width=600,height=260,resizable=yes,scrollbars=yes,status=yes');
			break;
		case 3:
			// 2009/03/10 R.Shibuta Upd-S
			// Win = window.open('sst300.asp', 'FConIn', 'width=710,height=650,resizable=yes,scrollbars=yes,status=yes');
			Win = window.open('sst300.asp', 'FConIn', 'width=900,height=900,resizable=yes,scrollbars=yes,status=yes');
			// 2009/03/10 R.Shibuta Upd-E -->
			break;
	}
	if(i==1){
		Fname.submit();
	}
}
-->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY class="menu">
<!-------------ステータス配信初期画面Menu--------------------------->
<CENTER>
<P><BR></P>
<P><B><Font color="#000066">DISPLAY SWITCH</FONT></B></P>
<P><A HREF="JavaScript:opnewin(1)">REQUESTING LIST</A></P>
<P><A HREF="JavaScript:opnewin(2)">INITIAL REQUEST</A></P>
<P><A HREF="JavaScript:opnewin(3)">SET UP</A></P>
<FORM name="sst000MF">
</FORM>
</CENTER>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
