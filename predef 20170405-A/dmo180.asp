<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo180.asp				_/
'_/	Function	:空搬入情報一覧CSVファイルダウンロード	_/
'_/	Date		:2003/07/31				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="./Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b208", "空搬入事前情報CSVファイルダウンロード","01",""


'データ取得
  dim Num,DtTbl,i,j
  Get_Data Num,DtTbl

' ファイルのダウンロード
  Response.ContentType="application/octet-stream"
  Response.AddHeader "Content-Disposition","attachment; filename=output.csv"
    Response.Write "搬入予定日,指示元,指示元への回答,コンテナ番号,船社,船名,サイズ,返却先,"
    Response.Write "ディテンションフリータイム,指示先,指示先回答,備考"
    Response.Write Chr(13) & Chr(10)
    For j=1 To Num
      Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(8))&","&Trim(DtTbl(j)(3))&","
      Response.Write DtTbl(j)(9)&","&Trim(DtTbl(j)(10))&","&Trim(DtTbl(j)(11))&","&Trim(DtTbl(j)(12))&","
      Response.Write Trim(DtTbl(j)(13))&","&Trim(DtTbl(j)(5))&","&Trim(DtTbl(j)(6))&","&Trim(DtTbl(j)(14))&","
      Response.Write Chr(13) & Chr(10)
    Next
  Response.End

%>
