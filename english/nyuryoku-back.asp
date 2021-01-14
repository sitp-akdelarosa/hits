<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-te.asp"

	' useridのセッションを空にする
	Session.Contents("userid") = ""

    ' ユーザーID入力画面へ
    Response.Redirect "userchk.asp?link=nyuryoku-te.asp"
%>
