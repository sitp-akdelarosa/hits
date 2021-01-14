<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' 指定引数の取得（行き先画面情報）
    Dim strLinkID
    strLinkID = Request.QueryString("link")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    Select Case strLinkID
        Case "http://www.fk-tosikou.or.jp"  strLinkNamne = "福岡北九州高速道路公社"
        Case "http://www.jartic.or.jp"      strLinkNamne = "（財）日本道路交通情報センター"
        Case Else                           strLinkNamne = "不明"
    End Select

    ' リンク情報を出力
	strLogInfo = "ゲート前映像・混雑状況紹介-リンク-" & strLinkNamne
    WriteLog fs, "8001",strLogInfo,"01", ","

    ' 行き先画面へリダイレクト
    Response.Redirect strLinkID
%>
