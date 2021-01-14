<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<%
' 海陸一貫システム(携帯電話版)　遷移図
' 
' 　　　　　　　　　コンテナNo照会
' 　　　　　　　　┌──→(mcont01.asp) ─→(mcont02.asp)
' 　　　　　　　　│　　　コンテナNo入力　　　照会結果
' 　　　　　　　　│　　　　　　　　　　　　　　↑
' 　　　　　　　　│BL番号照会　　　　　　　　　│
' メインメニュー　├──→(mblno01.asp) ─→(mblno02.asp)
'  (index.asp)　─┤　　　　BL番号入力　　　コンテナ一覧
' 　　　　　　　　│
' 　　　　　　　　│完了時刻入力
' 　　　　　　　　├──→(muser.asp) ─→(mrung01.asp) ─→(mrung02.asp) ─→[mrung03.asp]
' 　　　　　　　　│　　　ユーザID入力　　コンテナNo入力　　完了作業選択　　　ファイル出力
' 　　　　　　　　│
' 　　　　　　　　│ゲート内所用時間
' 　　　　　　　　├──→(mterm01.asp)
' 　　　　　　　　│
' 　　　　　　　　│映像表示(かもめ大橋,待機場,ゲート前 各共通)
' 　　　　　　　　└──→(mpict01.asp)
' 
%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "0200", "携帯-ＴＯＰ画面", "00" , sPhoneType, ","
Set oFS = Nothing

Dim sTBorder
If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="HiTS ver.2">
		<center>
		<img src="mtitle<%=GetImageExt()%>" alt="HiTS ver.2"><br>
		ｺﾝﾃﾅ搬出許可照会<br>
		<a task="gosub" dest="mcont01.asp" accesskey="1">　ｺﾝﾃﾅ番号照会</a><br>
		<a task="gosub" dest="./cam/mcont01cam.asp">　　　　中央ふ頭</a><br>
		<a task="gosub" dest="mblno01.asp" accesskey="2">　BL番号照会</a><br>
		<a task="gosub" dest="./cam/mblno01cam.asp">　　　　中央ふ頭</a><br>
		運行情報入力<br>
		<a task="gosub" dest="muser.asp" accesskey="3">　完了時刻入力</a><br>
		ターミナル情報<br>
		<a task="gosub" dest="mterm01.asp" accesskey="4">　ｹﾞ-ﾄ内時間</a><br>
		(香椎)<br>
		<a task="gosub" dest="mpict01.asp?pict=1" accesskey="5">　かもめ大橋</a><br>
		<a task="gosub" dest="mpict01.asp?pict=2" accesskey="6">　待機場映像</a><br>
		<a task="gosub" dest="mpict01.asp?pict=3" accesskey="7">　ｹﾞ-ﾄ前映像</a><br>
		(ICCT)<br>
		<a task="gosub" dest="mpict01.asp?pict=4" accesskey="8">　ｹﾞ-ﾄ前映像</a><br>
	</display>
	</hdml>
<%
Else
	' EzWeb以外のタグを編集
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("HiTS ver.2")%>
	</head>
	
	<body>
	<center>
	<img src="mtitle<%=GetImageExt()%>" alt="HiTS ver.2"><br>
	<hr>
<%
	If sPhoneType = "I" Then
		'i-modeではテーブルがつかえないので左詰め
%>
		<div align="left">
			ｺﾝﾃﾅ搬出許可照会<br>
			　<a href="mcont01.asp" <%=GetKeyTag("1")%>><%=GetKeyLabel("1")%>ｺﾝﾃﾅ番号照会</a><br>
			　<a href="mblno01.asp" <%=GetKeyTag("2")%>><%=GetKeyLabel("2")%>BL番号照会</a><br>
			運行情報入力<br>
			　<a href="muser.asp" <%=GetKeyTag("3")%>><%=GetKeyLabel("3")%>完了時刻入力</a><br>
			ターミナル情報<br>
			　<a href="mterm01.asp" <%=GetKeyTag("4")%>><%=GetKeyLabel("4")%>ｹﾞ-ﾄ内時間</a><br>
			　(香椎)<br>
			　<a href="mpict01.asp?pict=1" <%=GetKeyTag("5")%>><%=GetKeyLabel("5")%>かもめ大橋</a><br>
			　<a href="mpict01.asp?pict=2" <%=GetKeyTag("6")%>><%=GetKeyLabel("6")%>待機場映像</a><br>
			　<a href="mpict01.asp?pict=3" <%=GetKeyTag("7")%>><%=GetKeyLabel("7")%>ｹﾞ-ﾄ前映像</a><br>
			　(ICCT)<br>
			　<a href="mpict01.asp?pict=4" <%=GetKeyTag("8")%>><%=GetKeyLabel("8")%>ｹﾞ-ﾄ前映像</a><br>
		</div>
<%
	Else
		If sPhoneType = "J" Then
			sTBorder = ""
		Else
			sTBorder = " border=""0"" "
		End If
%>
		<table <%=sTBorder%>>
			<tr><td>
				ｺﾝﾃﾅ搬出許可照会<br>
			</td></tr>
			<tr><td>
				　<a href="mcont01.asp" <%=GetKeyTag("1")%>><%=GetKeyLabel("1")%>ｺﾝﾃﾅ番号照会</a><br>
			</td></tr>
			<tr><td>
				　<a href="mblno01.asp" <%=GetKeyTag("2")%>><%=GetKeyLabel("2")%>BL番号照会</a><br>
			</td></tr>
			<tr><td>
				運行情報入力<br>
			</td></tr>
			<tr><td>
				　<a href="muser.asp" <%=GetKeyTag("3")%>><%=GetKeyLabel("3")%>完了時刻入力</a><br>
			</td></tr>
			<tr><td>
				ターミナル情報<br>
			</td></tr>
			<tr><td>
				　<a href="mterm01.asp" <%=GetKeyTag("4")%>><%=GetKeyLabel("4")%>ｹﾞ-ﾄ内時間</a><br>
			</td></tr>
			<tr><td>
				　(香椎)<br>
			</td></tr>
			<tr><td>
				　<a href="mpict01.asp?pict=1" <%=GetKeyTag("5")%>><%=GetKeyLabel("5")%>かもめ大橋</a><br>
			</td></tr>
			<tr><td>
				　<a href="mpict01.asp?pict=2" <%=GetKeyTag("6")%>><%=GetKeyLabel("6")%>待機場映像</a><br>
			</td></tr>
			<tr><td>
				　<a href="mpict01.asp?pict=3" <%=GetKeyTag("7")%>><%=GetKeyLabel("7")%>ｹﾞ-ﾄ前映像</a><br>
			</td></tr>
			<tr><td>
				　(ICCT)<br>
			</td></tr>
			<tr><td>
				　<a href="mpict01.asp?pict=4" <%=GetKeyTag("8")%>><%=GetKeyLabel("8")%>ｹﾞ-ﾄ前映像</a><br>
			</td></tr>
		</table>
<%
	End If
%>
	<hr>
	</body>
	</html>
<%
End If
%>




