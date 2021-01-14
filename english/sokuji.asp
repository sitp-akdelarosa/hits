<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "sokuji.asp"

    ' ユーザ種類を取得する
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"             'トップ
        Response.End
    End If

    ' ユーザ種類により画面を選択
    If strUserKind="海貨" Then
        Response.Redirect "sokuji-kaika-updtchk.asp"
    Else
        Response.Redirect "sokuji-koun-updtchk.asp"
    End If
    Response.End
%>
