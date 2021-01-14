<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' 指定引数の取得（行き先画面情報）
    Dim strLinkUrl
    Dim strLogId
    Dim strLogNo
    Dim strLinkName

    strLogId = Request.QueryString("longid")
    strLogNo = Request.QueryString("logno")
    strLinkName = Request.QueryString("linkname")
    strLinkUrl = Request.QueryString("linkurl")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    If strLogId = "" or strLogNo = "" or strLinkName = "" Then
        strLogId = "l999"
        strLogNo = "01"
        strLinkName = "その他"
    End If

    ' リンク情報を出力
    WriteLog fs, strLogId, strLinkName, strLogNo, ","

    ' 行き先画面へリダイレクト
    Response.Redirect strLinkUrl
%>
