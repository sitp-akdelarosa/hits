<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst200.asp				_/
'_/	Function	:ステータス配信依頼登録画面			_/
'_/	Date			:2004/01/07				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''セッションの有効性をチェック
	CheckLoginH

	'''データ取得
	Dim USER
	USER = Session.Contents("userid")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信依頼登録</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
	window.resizeTo(450,180);
	bgset(target);
	window.focus();
}

function GoNext(){
	f=document.sst200;
	Number=LTrim(f.ContBLNo.value);
	if(Number.length==0){
		alert("コンテナ番号またはＢＬ番号を入力してください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(Number)){
		alert("コンテナ番号に半角英数字以外の文字を入力しないでください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && Number.length>12){
		alert("コンテナ番号は１２文字以内で指定してください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && !CheckEisu(Number)){
		alert("ＢＬ番号に半角英数字と半角スペース、「-」、「/」以外の文字を入力しないでください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && Number.toUpperCase() == "EMPTY"){
		alert("EMPTYは登録できません。");
		f.ContBLNo.focus();
		return;
	}

	changeUpper(f);
	f.action="sst201.asp";
	f.submit();
}

function GoSendmail(){
	f=document.sst200;
	Number=LTrim(f.ContBLNo.value);
	if(Number.length==0){
		alert("コンテナ番号またはＢＬ番号を入力してください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(Number)){
		alert("コンテナ番号に半角英数字以外の文字を入力しないでください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && Number.length>12){
		alert("コンテナ番号は１２文字以内で指定してください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && !CheckEisu(Number)){
		alert("ＢＬ番号に半角英数字と半角スペース、「-」、「/」以外の文字を入力しないでください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && Number.toUpperCase() == "EMPTY"){
		alert("EMPTYは指定できません。");
		f.ContBLNo.focus();
		return;
	}

	if(!confirm("送信してもよろしいですか？")){
		f.ContBLNo.focus();
		return;
	}
	f.Mode.value=1;		//新規登録画面よりmail即時送信を実行した場合
	f.action="sst500.asp";
	f.submit();
}
//コンテナ情報照会
function GoInfo(){
	f=document.sst200;
	Number=LTrim(f.ContBLNo.value);
	if(Number.length==0){
		alert("コンテナ番号またはＢＬ番号を入力してください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(Number)){
		alert("コンテナ番号に半角英数字以外の文字を入力しないでください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && Number.length>12){
		alert("コンテナ番号は１２文字以内で指定してください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && !CheckEisu(Number)){
		alert("ＢＬ番号に半角英数字と半角スペース、「-」、「/」以外の文字を入力しないでください。");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && Number.toUpperCase() == "EMPTY"){
		alert("EMPTYは指定できません。");
		f.ContBLNo.focus();
		return;
	}

	f.action="sst900.asp";
	newWin = window.open("", "ConInfo", "status=yes,scrollbars=yes,resizable=yes,menubar=yes");
	f.target="ConInfo";
	f.submit();
	f.target="_self";
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0"  onLoad="setParam(document.sst200)">
<!-------------ステータス配信依頼登録画面--------------------------->
<% Session.Contents("InsertSubmitted")="False"  %>
<% Session.Contents("SendMailSubmitted")="False"  %>
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst200" method="POST">
	<TR>
		<TD colspan="3">
			<B>輸入ステータス配信依頼登録</B><BR>
		</TD>
	</TR>
	<TR>
		<TD width="40%"><DIV class="bgb">ログインユーザ</DIV></TD>
		<TD width="60%" colspan="2">
			<INPUT type="text" name="LoginUser" value="<%=USER%>" size="10" readonly style="background-color:#E0E0E0;color:#000000;">
		</TD>
	</TR>
	<TR>
		<TD width="40%"><DIV class="bgb">対象コンテナNo.／ＢＬNo.</DIV></TD>
		<TD width="40%">
			<INPUT type="text" name="ContBLNo" value="" size="27" maxlength="20">
		</TD>
		<TD width="20%">
			<INPUT type="button" value="情報照会" onClick="GoInfo()">
		</TD>
	</TR>
	<TR>
		<TD colspan="3" align="center">
			<INPUT type="radio" name="ContORBL" value="1" checked>コンテナ　
			<INPUT type="radio" name="ContORBL" value="2">ＢＬ
		</TD>
	</TR>
	<TR>
		<TD colspan="3" align="center">
			<INPUT type="hidden" name="Mode" value="">
			<INPUT type="button" value="登録" onClick="GoNext()">
			<INPUT type="button" value="中止" onClick="window.close()">　
			<A HREF="javascript:GoSendmail();">mail即時送信</A>
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
