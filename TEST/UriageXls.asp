<%@ LANGUAGE="VbScript" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<!--#include file="XlsCrt3vbs.inc"-->
<HTML>
<HEAD>
<TITLE>売上伝票ファイル出力</TITLE>
</HEAD>
<BODY>
<%
	On Error Resume Next

	wDate = FormatDateTime (Date,vbLongDate)
	wTime = FormatDateTime (Time,vbLongTime)
	wIPAddress    = Request.ServerVariables("REMOTE_ADDR")

	wOutFileName = "Xls3_Uriage" & wDate & Replace(wTime,":","_") & wIPAddress & ".xls"

	'---------------------------------------------------------
	' 以下のwFilePath、wIISFilePath、wInFileNameの値は、
        ' 実行環境の構成に従い変更して下さい
	'---------------------------------------------------------
	wFilePath  = "c:\inetpub\wwwroot\outfiles\"
	wIISFilePath = "http://localhost/outfiles/"
	wInFileName = "c:\Xls3Sample.xls"

	

	'--------------------------------------------------------
	'　入力内容を変数に取得
	'--------------------------------------------------------
	wUDate    =  Request("UDate")    '売上日
	wUNo      =  Request("UNo")      '伝票No
	wTName    =  Request("TName")    '得意先名
	wTAddress =  Request("TAddress") '得意先住所（納品先）
	wShimei   =  Request("Shimei")   '氏名（納品先）
	wSCode    =  Request("SCode")    '商品コード
	wSName    =  Request("SName")    '商品名
	wSuu      =  Request("Suu")      '数量
	wTanka    =  Request("Tanka")    '単価
	
	'--------------------------------------------------------
	'  ExcelCreator オブジェクト生成→Excelファイル出力
	'--------------------------------------------------------
        'ExcelCreator オブジェクト生成
        Set Xls1= Server.CreateObject("ExcelCrtOcx.ExcelCrtOcx.1")  

	'売上伝票(オーバーレイ)ファイルオープン
  	Xls1.OpenBook wFilePath & wOutFileName,wInFileName

        '雛型シートを呼び出し
        Xls1.SheetNo = 0

    'ブラウザ上で入力したデータをシートに出力
	Xls1.Cell("**UDate").Str    = wUDate         '売上日
	Xls1.Cell("**UNo").Str      = wUNo           '伝票No
	Xls1.Cell("**TName").Str    = wTName & "様"  '得意先名
	Xls1.Cell("**TAddress").Str = wTAddress      '得意先住所（納品先）
	Xls1.Cell("**Shimei").Str   = wShimei & "様" '氏名（納品先）
	Xls1.Cell("**SCode").Str    = wSCode         '商品コード
	Xls1.Cell("**SName").Str    = wSName         '商品名
	Xls1.Cell("**Suu").Long     = CLng(wSuu)     '数量
	Xls1.Cell("**Tanka").Double = CDbl(wTanka)   '単価


	wGoukei = CLng(wSuu) * CDbl(wTanka) '合計金額（税抜き）
	Xls1.Cell("**Kingaku").Double = CDbl(wGoukei)

	wZei = wGoukei * 100 * 0.05 / 100  '税額
	Xls1.Cell("**Zei").Value = wZei
	
	Xls1.Cell("I18").Func2 "=SUM(C18,F18)",wGoukei + wZei   '税込み合計金額欄

	wMsg = "Excelファイルを作成しました。以下より作成したファイルをダウンロードできます"

	wErrNo = Xls1.ErrorNo
	If wErrNo <> 0 Then
		wMsg = "ExcelCreator3エラーメッセージ：" & Xls1.ErrorMessage
	End If
	Xls1.CloseBook

        Set Xls1 = Nothing
%>
<FONT SIZE="2"><%=wMsg%></FONT><BR>
<% If wErrNo = 0 Then %>
    <Font Size="2">生成したファイルのダウンロード</font><br>
    <Font Size="2"><a href="<%=wIISFilePath%><%=wOutFileName%>"><%=wOutFileName%></A></font>
<% End If %>
</BODY>
</HTML>