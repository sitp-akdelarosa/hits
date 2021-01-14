<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst221.asp				_/
'_/	Function	:ステータス配信対象削除			_/
'_/	Date			:2004/01/15				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''セッションの有効性をチェック
	CheckLoginH

	'''データ削除しました画面にて「最新の情報に更新」でSubmitされた場合の対策
	if Session.Contents("DeleteSubmitted")="False" then

	'''データ取得
	Dim USER, KIND, NUMBER
	USER   = UCase(Session.Contents("userid"))
	KIND = Request.Form("ContORBL")
	NUMBER = Request.Form("ContBLNo")

	'''エラートラップ開始
	on error resume next
	'''DB接続
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	'''データ削除（処理区分Processを'D'にする。実際のレコード削除は日次処理にて行う。）
	StrSQL = "UPDATE TargetContainers SET UpdtTime='" & Now() & "', UpdtPgCd='STATUS01',"
	StrSQL =  StrSQL & " UpdtTmnl='" & USER & "', Process='D' "
	if KIND = 1 then		'''削除対象がコンテナ番号
		StrSQL =  StrSQL & " WHERE ContNo='" & NUMBER & "' AND UserCode='" & USER & "'"
	elseif KIND = 2 then		'''削除対象がＢＬ番号
		StrSQL =  StrSQL & " WHERE BLNo='" & NUMBER & "' AND UserCode='" & USER & "'"
	else
		response.write("KIND error!")
		response.end
	end if

	ObjConn.Execute(StrSQL)
	if err <> 0 then
		Set ObjRS = Nothing
		jumpErrorPDB ObjConn,"1","c102","14","ステータス配信対象削除","104","SQL:<BR>"&StrSQL
	end if

	'''ログ出力
	WriteLogH "c102", "ステータス配信対象削除","14",""
	ObjRS.close

	'''DB接続解除
	DisConnDBH ObjConn, ObjRS
	'''エラートラップ解除
	on error goto 0

	Session.Contents("DeleteSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信対象削除</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信対象削除--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst221">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			削除しました。
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("DeleteSubmitted")="False"のelse処理 %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信対象削除</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信対象削除--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst221">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			削除は既に完了しています。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("DeleteSubmitted")="False"のendif処理 %>
<% end if %>
