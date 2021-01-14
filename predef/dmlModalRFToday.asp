<% @LANGUAGE = VBScript %>
<%
%><% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
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
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="">
<form name="frm" method="post">

<table border=0 cellPadding=1 cellSpacing=0 width="100%">
<tr>
<td align=center>
<TABLE border=0 cellPadding=4 cellSpacing=0>
  <tr>
  <td colspan=2 align=center>
	<div><BR></div>
    <div align=left style="font-size:15px;">当日の搬出予約となっております。<BR><BR>
         コンテナの貸出準備が間に合わない可能性がございますので、<BR><BR>
         担当オペレータまで確認連絡をお願いします。
    </div>
  </td>
  </tr>
  <tr><td><BR /></td></tr>
  <tr>
  <td align=center><input type="button" name="Back" value="中止" Onclick="fBack();" onkeypress="return true" style="font-size:15px;"></td>
  <td align=center><input type="button" name="Rgst" value="継続" Onclick="fRgst();" onkeypress="return true" style="font-size:15px;"></td>
  </tr>
</TABLE>
</td>
</tr>
</table>
</div>
</form>
</BODY>
</HTML>
