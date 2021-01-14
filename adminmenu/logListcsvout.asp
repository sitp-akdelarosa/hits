<%@Language="VBScript" %>

<!--#include file="./Common/Common.inc"-->

<%
	'変数宣言
	Dim strFileName

	' Tempファイル属性のチェック

	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' ダウンロードファイルの取得

	strFileName = Session.Contents("tempfile")
	If strFileName="" Then
		' セッションが切れているとき
		Response.Redirect "accesstotal.asp"	 '利用件数Topへ
		Response.End
	End If
	strFileName="../temp/" & strFileName
	' ダウンロードファイルのOpen
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' ファイルのダウンロード
	Response.ContentType="application/octet-stream"
	Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

	'ヘッダ書き込み
	Response.Write "アクセス件数累計表"
	Response.Write Chr(13) & Chr(10)
	Response.Write Chr(13) & Chr(10)

	'CSVファイルタイトル行出力
	Response.Write "区分,"
	Response.Write "PC,"
	Response.Write "携帯端末,"
	'Y.TAKAKUWA Add-S 2013-09-30
	Response.Write "形態端末,"
	'Y.TAKAKUWA Add-E 2013-09-30
	Response.Write "合計," 
	Response.Write "累計" 
	Response.Write Chr(13) & Chr(10)
	
	'累計CSVファイルデータ行出力
	Do While Not ti.AtEndOfStream
		anyTmp=Split(ti.ReadLine,",")
		Response.Write anyTmp(0) & ","
		Response.Write anyTmp(1) & ","
		Response.Write anyTmp(2) & ","
		'2013-09-30 Y.TAKAKUWA Upd-S
		'Response.Write anyTmp(3) & ","
		'Response.Write anyTmp(4) & ""
		Response.Write anyTmp(3) & ","
		Response.Write anyTmp(4) & ","
		Response.Write anyTmp(5) & ""
		'2013-09-30 Y.TAKAKUWA Upd-E
		Response.Write Chr(13) & Chr(10)
	Loop


	' ダウンロード終了
	Response.End

%>
