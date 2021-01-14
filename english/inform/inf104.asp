<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf104.asp                                      _/
'_/	Function	:お知らせ送信先情報の削除処理                    _/
'_/	Date			:2005/03/10                                      _/
'_/	Code By		:aspLand HARA                                    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''削除更新しました画面にて「最新の情報に更新」でSubmitされた場合の対策
	if Session.Contents("DeleteSubmitted")="False" then

		'''データ取得
		Dim EMAIL
		EMAIL = Request.Form("email")

		'''エラートラップ開始
		on error resume next
		'''DB接続
		Dim cn, rs, sql
		ConnDBH cn, rs

		sql="delete from send_information where email='" & EMAIL & "'"

		cn.Execute(sql)
		if err <> 0 then
			set rs = Nothing
			response.write("inf104.asp:send_informationテーブルdeleteエラー!")
			response.end
		end if

		'''DB接続解除
		DisConnDBH cn, rs
		'''エラートラップ解除
		on error goto 0

		Session.Contents("DeleteSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>お知らせ送信先情報削除</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="inf101.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------お知らせ送信先情報削除--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf103">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			削除しました。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("UpdateSubmitted")="False"のelse処理 %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>お知らせ送信先情報削除</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="inf101.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------お知らせ送信先情報削除--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf103">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			削除はすでに完了しています。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("UpdateSubmitted")="False"のendif処理 %>
<% end if %>
