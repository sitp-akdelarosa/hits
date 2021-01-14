<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst900.asp				_/
'_/	Function	:ステータス情報配信共通処理			_/
'_/	Date			:2004/1/15				_/
'_/	Code By		:aspLand HARA		_/
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
	Dim CONnum,Flag,BLnum
	Dim inPutStr,strNums
	CONnum = Request.Form("ContBLNo")
	Flag   = Request.Form("ContORBL")

	'''エラートラップ開始
	on error resume next
	'''DB接続
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	Select Case Flag
		Case "1"		'''コンテナ番号指定
			inPutStr="<INPUT type=hidden name='cntnrno' value='"& CONnum &"'>"
		Case "2"		'ＢＬ番号指定
			inPutStr="<INPUT type=hidden name='blno' value='"& CONnum &"'>"
	End Select

	if Flag=1 Then
		Session.Contents("route") = "輸入コンテナ情報照会（作業選択） "
	Else
		Session.Contents("route") = "Top > 輸入コンテナ情報照会（作業選択） "
	End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>転送中</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function opnewin(){
  window.focus();
  document.sst900.submit();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="opnewin()">
<P>転送中...しばらくお待ちください。</P>
<FORM action="../impcntnr.asp" name="sst900">
<%= inPutStr %>
</FORM>
</BODY>
</HTML>
