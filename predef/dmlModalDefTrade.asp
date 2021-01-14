<% @LANGUAGE = VBScript %>
<%
%><% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="CommonFunc.inc"-->
<!--#include File="Common.inc"-->
<%
  Dim ObjRS2,ObjConn2
  Dim StrSQL
  Dim DefCodeVal
  Dim DefNameVal

  DefCodeVal = gfTrim(Request.QueryString("DfTCd"))
  DefNameVal = ""

  'セッションの有効性をチェック
  CheckLoginH
    
  if gfTrim(DefCodeVal) <> "" then
	  '確定事業者取得
	  ConnDBH ObjConn2, ObjRS2

	  StrSQL = "SELECT DefName FROM mDefTrade "
	  StrSQL = StrSQL & " WHERE DefCode = '" & gfSQLEncode(DefCodeVal) & "'"
	  ObjRS2.Open StrSQL, ObjConn2

	  if err <> 0 then
		DisConnDBH ObjConn2, ObjRS2	'DB切断
		jampErrerP "2","b301","01","船名・次航検索","102","SQL:<BR>" & StrSQL & err.description & Err.number
	  end if
	  
	  if not ObjRS2.EOF then
	    DefNameVal = gfTrim(ObjRS2("DefName"))
	  end if
	  
	  ObjRS2.close
	  DisConnDBH ObjConn2, ObjRS2	'DB切断
  end if


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function fSend()
{
<%
  Response.Write "window.returnValue = '" & DefNameVal & "';"
%>
  window.close();
}

-->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="fSend();">
<form name="frm" method="post">

<table width="100%" height="82%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="50" nowrap>&nbsp;</td>
<td>
  <div id="BDIV3" style="width: 100%; height: 100%; padding-top:20px;">
  <input type="hidden" name="DefCodeVal" value="<%=DefCodeVal%>"/>
  <input type="hidden" name="DefNameVal" value="<%=DefNameVal%>"/>
  </div>
</td></tr>  
</table>
</form>
</BODY>
</HTML>
