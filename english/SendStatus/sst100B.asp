<%@LANGUAGE = VBScript%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst100B.asp				_/
'_/	Function	:ステータス配信依頼中一覧画面フッタ		_/
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
<TITLE>IMPORT STATUS DELIVERY REQUEST LIST</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信依頼中一覧画面Bottom--------------------------->
<CENTER>
	<FORM name="next" action="">
	<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" height=35>
		<TR>
			<TD>
				<A HREF="JavaScript:GoHelp(1)">HELP</A>・・・Click here for function introduction. 
			</TD>
		</TR>
		<INPUT type=hidden name="SortFlag" value="">
	</TABLE>
	</FORM>
</CENTER>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
