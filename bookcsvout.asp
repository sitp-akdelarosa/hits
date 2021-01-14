<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "EXPORT", "bookentry.asp"

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ダウンロードファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "bookentry.asp"             '輸出コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' ダウンロードファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' ファイルのダウンロード
    Response.ContentType="application/octet-stream"
    Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

    '輸出コンテナCSVファイルタイトル行出力
    Response.Write "Booking No.,"
    Response.Write "船社,"
    Response.Write "船名,"
    Response.Write "Voyage No.,"
    Response.Write "仕向港,"
    Response.Write "空コン搬出場所,"
    Response.Write "CYカット,"	'I20080222
    Response.Write "サイズ,"
    Response.Write "タイプ,"
    Response.Write "高さ,"
    Response.Write "材質,"	'I20040223
    Response.Write "予約本数,"
    Response.Write "搬出済本数"

    Response.Write Chr(13) & Chr(10)

    '輸出コンテナCSVファイルデータ行出力
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")

        Response.Write anyTmp(1) & ","
        Response.Write anyTmp(2) & ","
        Response.Write anyTmp(3) & ","
        Response.Write anyTmp(4) & ","
        Response.Write anyTmp(5) & ","
        Response.Write anyTmp(6) & ","
        Response.Write anyTmp(14) & ","		'I20080222
        Response.Write anyTmp(7) & ","
        Response.Write anyTmp(8) & ","
        Response.Write anyTmp(9) & ","
        Response.Write anyTmp(12) & ","		'I20040223
        Response.Write anyTmp(10) & ","
        Response.Write anyTmp(11)

'		If UBound(anyTmp)>11 Then
''			For i=12 To UBound(anyTmp)	'D20040223
'			For i=13 To UBound(anyTmp)	'I20040223
'				Response.Write "," & anyTmp(i)
'			Next
'		End If

        Response.Write Chr(13) & Chr(10)
    Loop

   ' 輸出コンテナ照会
    WriteLog fs, "1013","ブッキング情報照会-CSVファイル出力","30", ","

    ' ダウンロード終了
    Response.End

%>
