<% @LANGUAGE = VBScript %>
<%
%><% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
  Dim ChkAgree
  Dim ChkSolas
  Dim BookChk
  Dim IMDGChk										'2016/11/03 H.Yoshikawa Add
  ChkAgree = Trim(Request.QueryString("ChkAgr"))
  ChkSolas = Trim(Request.QueryString("ChkSls"))
  BookChk = Trim(Request.QueryString("BookChk"))
  IMDGChk = Trim(Request.QueryString("IMDGChk"))	'2016/11/03 H.Yoshikawa Add
'セッションの有効性をチェック
  CheckLoginH
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function finit()
{
    if("<%=ChkAgree%>" == "1" && "<%=ChkSolas%>" == "1" && "<%=BookChk%>" == "0"  && "<%=IMDGChk%>" == "0" ){		// 2016/11/03 H.Yoshikawa Upd(IMDGChk追加)
        fRgst();
    }
}
function fBack()
{
   returnValue = false;
   window.close();
}
function fRgst()
{
  returnValue = true;
  window.close();
}
-->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="finit()">
<form name="frm" method="post">

<table border=0 cellPadding=1 cellSpacing=0 width="100%">
<tr>
<td align=center>
<TABLE border=0 cellPadding=4 cellSpacing=0>
  <tr>
  <td colspan=2 align=center>
	<div><BR></div>
    <div>この登録は、下記の理由により「仮登録」状態となり、ゲート受付はできません。</div>
	<div><BR></div>
	<div><BR></div>
  <% If Trim(ChkSolas) <> "1" Then %>
	<div align=left>■「ここに入力したコンテナグロスはSOLAS条約に基づく方法で計測された数値です。」に<BR>　チェックがありません。</div>
  <%End If%>
  <% If Trim(ChkAgree) <> "1" Then %>
	<div align=left>■「本画面の入力内容をゲートでの搬入票の代わりとして使用することに同意します。」に<BR>　チェックがありません。</div>
  <%End If%>
  <% If Trim(BookChk) <> "0" Then %>
	<div align=left>■ブッキング情報と異なる値があります。（赤色表示）</div>
  <%End If%>
  <% '2016/11/03 H.Yoshikawa Add Start %>
  <% If Trim(IMDGChk) <> "0" Then %>
	<div align=left>■危険品コードに誤りがあります。（赤色表示）</div>
  <%End If%>
  <% '2016/11/03 H.Yoshikawa Add End %>
	<div><BR></div>
	<div><BR></div>
    <div>本登録を行う場合は、「戻る」ボタンを押して、下記の値を修正のうえご登録ください。</div>
    <div>このまま「仮登録」を行いますか？</div>
  </td>
  </tr>
  <tr><td><BR /></td></tr>
  <tr>
  <td align=center><input type="button" name="Back" value="戻る" Onclick="fBack();" onkeypress="return true"></td>
  <td align=center><input type="button" name="Rgst" value="仮登録" Onclick="fRgst();" onkeypress="return true"></td>
  </tr>
</TABLE>
</td>
</tr>
</table>
</div>
</form>
</BODY>
</HTML>
