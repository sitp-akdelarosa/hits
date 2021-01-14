<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst100T.asp				_/
'_/	Function	:ステータス配信依頼中一覧画面トップ		_/
'_/	Date			:2003/12/25				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'ユーザデータ所得
	Dim USER, COMPcd, LinUN
	USER   = Session.Contents("userid")
	COMPcd = Session.Contents("COMPcd")
	LinUN  = Session.Contents("LinUN")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信依頼中一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
//<!--
//検索
function SearchContBL(){
	f = document.search;
	if(!f.ContORBL[0].checked && !f.ContORBL[1].checked){
		alert("検索対象を選択してください");
		return false;
	}
	Num=LTrim(f.ContBLNo.value);
	if(Num.length==0){
		alert("検索する番号を入力してください");
		f.ContBLNo.focus();
		return false;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(f.ContBLNo.value)){
		alert("検索するコンテナ番号に半角英数字以外の文字を指定しないでください");
		f.ContBLNo.focus();
		return false;
	}
	if(f.ContORBL[1].checked && !CheckEisu(f.ContBLNo.value)){
		alert("検索するコンテナ番号に半角英数字と半角スペース、「-」、「/」以外の文字を指定しないでください");
		f.ContBLNo.focus();
		return false;
	}
	if(f.ContORBL[0].checked){
		parent.DList.SearchC("2",f.ContBLNo.value);
	} else if(f.ContORBL[1].checked){
		parent.DList.SearchC("3",f.ContBLNo.value);
	}
}
//ソート
//function sort(){
//	f = document.search;
//	f.SortFlag.value=f.Sort.options[target.Sort.selectedIndex].value;
//	f.target="DList";
//	f.action="./sst100L.asp";
//	f.submit();
//}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信依頼中一覧画面Top--------------------------->
<TABLE border="1" cellPadding="3" cellSpacing="0" align=center width="100%">
	<TR>
		<TD width="15%">ログインユーザ</TD>
		<TD><%=LinUN%></TD>
	</TR>
</TABLE>
<TABLE border="0" cellPadding="3" cellSpacing="0" width="100%">
	<FORM name="search" action="">
	<TR>
		<TD width="50%"><BR><B class=title>ステータス配信依頼中一覧</B></TD>
		<TD width="50%">
			<INPUT type="hidden" name="ContBLFlag" value="">
			<INPUT type="radio" name="ContORBL">コンテナ番号
			<INPUT type="radio" name="ContORBL">ＢＬ番号<BR>
			<INPUT type="text"  name="ContBLNo" maxlength="20">
			<INPUT type="button" value="検索" onClick="SearchContBL()">
		</TD>
	<TR>
	</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
