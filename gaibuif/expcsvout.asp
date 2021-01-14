<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "EXPORT", "expentry.asp"

    ' 表示モードの取得
    Dim bDispMode          ' true=コンテナ検索 / false=ブッキング検索
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
        strOption = "コンテナNo.CSVファイル送信"
    Else
        bDispMode = false
        strOption = "Booking番号CSVファイル送信"
    End If

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ダウンロードファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "expentry.asp"             '輸出コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' ダウンロードファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' ファイルのダウンロード
    Response.ContentType="application/octet-stream"
    Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

    '輸出コンテナCSVファイルタイトル行出力
    CsvTitleWrite bDispMode

    '輸出コンテナCSVファイルデータ行出力
    CsvDataWrite bDispMode, ti

    ' 輸出コンテナ照会
    WriteLog fs, "3008","仕出地情報照会-CSVファイル出力","30", ","

    ' ダウンロード終了
    Response.End

%>
