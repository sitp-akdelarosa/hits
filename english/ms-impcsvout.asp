<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="MS-ImpCom.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' ユーザ種類をチェックする
    strUserKind=Session.Contents("userkind")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ダウンロードファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "impentry.asp"             '輸入コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' ダウンロードファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' ファイルのダウンロード
    Response.ContentType="application/octet-stream"
    Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

    '輸入コンテナCSVファイルタイトル行出力
    CsvTitleWrite strUserKind

    '輸入コンテナCSVファイルデータ行出力
    CsvDataWrite strUserKind, ti

    ' 輸入コンテナ照会
    WriteLog fs, "2110","輸入コンテナ照会-CSVファイル出力","30", filename & ","

    ' ダウンロード終了
    Response.End

%>
