<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim vCtnoS, vCtnoE, vUserID
Dim sCntNo,sCntNo2
Dim sUserID
Dim sSQL
Dim sErrMsg
Dim sErrOpt

sErrMSg = ""
sErrOpt = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

vCtnoE = Trim(Request.QueryString("cont_e"))
vCtnoS = Trim(Request.QueryString("cont_s"))
vUserID = Trim(Request.QueryString("UserID"))

If (IsEmpty(vCtnoE) Or vCtnoE = "") And (IsEmpty(vCtnoS) Or vCtnoS = "") Then
	sErrMsg = "コンテナ未入力"
Else
	If IsEmpty(vUserID) Or vUserID = "" Then
		sErrMsg = "ユーザーID未入力"
	Else
		sUserID = vUserID
	End If
End If

If sErrMsg = "" Then
	Dim conn, rs
	ConnectSvr conn, rs

	'該当するコンテナを探す
	If IsEmpty(vCtnoE) Or vCtnoE = "" Then
		'コンテナ番号の数値部分のみ入力されている場合
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container GROUP BY RTrim([ContNo]), ContNo "
		sSQL = sSQL & "HAVING (((RTrim([ContNo])) Like '%" & vCtnoS & "'))"
	Else
		'コンテナ番号の英字部分、数値部分ともに入力されている場合
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container "
		sSQL = sSQL & "WHERE RTrim([ContNo]) = '" & UCase(vCtnoE) & vCtnoS & "'"
	End If
	rs.Open sSQL, conn, 0, 1, 1
	If rs.Eof Then
		sErrMsg = "該当コンテナなし"
		sErrOpt = vCtnoE & vCtnoS
	Else
		sCntNo = rs("CT")		'コンテナ番号再設定
		rs.MoveNext
		Do While Not rs.EOF
			sCntNo2 = rs("CT")
			rs.MoveNext
			If sCntNo<>sCntNo2 Then
				sErrMsg = "ｺﾝﾃﾅ複数存在"
				sErrOpt = vCtnoS
				Exit Do
			End If
		Loop
	End If
	rs.Close

	If sErrMsg = "" Then
		' 今回検索したコンテナ番号をユーザテーブルに保存(次回にデフォルトで表示する為)
		sSQL = "SELECT lUserTable.BeforeCntnrNo FROM lUserTable WHERE lUserTable.UserID='" & sUserID & "'"
		rs.Open sSQL, conn, 2, 2
		If Not rs.Eof Then
			rs("BeforeCntnrNo") = sCntNo
			rs.Update
		End If
		rs.Close
	End If

	conn.Close
End If

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If sErrMsg<>"" Then
	WriteLogM oFS, sUserID, "6201", "携帯-完了時刻コンテナ番号入力", "10",sPhoneType, vCtnoE & "/" & vCtnoS & "," & "入力内容の正誤:1(誤り)" & sErrMsg
Else
	WriteLogM oFS, sUserID, "6201", "携帯-完了時刻コンテナ番号入力", "10",sPhoneType, vCtnoE & "/" & vCtnoS & "," & "入力内容の正誤:0(正しい)"
	WriteLogM oFS, sUserID, "6202", "携帯-完了時刻入力", "00",sPhoneType, sCntNo & ","
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="運行情報入力">
		<center>
		【完了時刻入力】<br>
<%
		If sErrMsg <> "" Then
%>
			<center>
			<%=sErrOpt%><br>
			<center>
			<%=sErrMsg%><br><br>
			<center>
			<a task="gosub" dest="index.asp">ﾒﾆｭｰ</a>
<%
		Else
%>
			<center>
			コンテナ番号<br>
			<center>
			<%=sCntNo%><br>
			<center>
			完了した作業<br>
			<center>
			<a task="gosub" accesskey="3"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=C">出:空倉庫着
			</a><br>
			<center>
			<a task="gosub" accesskey="4"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=D">出:ﾊﾞﾝﾆﾝｸﾞ完
			</a><br>
			<center>
			<a task="gosub" accesskey="1"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=A">入:実入倉庫着
			</a><br>
			<center>
			<a task="gosub" accesskey="2"
				dest="mrung03.asp?UserID=<%=sUserID%>&Contno=<%=sCntNo%>&operation=B">入:デバン完
			</a><br>
<%
		End If
%>
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
		<%=GetTitleTag("運行情報入力")%>
	</head>
	
	<body>
	<center>
	【完了時刻入力】
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<%=sErrOpt%><br>
		<%=sErrMsg%><br><br>
		<form action="index.asp" method="get">
			<input type="submit" value="ﾒﾆｭｰ">
		</form>
<%
	Else
%>
		<form action="mrung03.asp" method="get">
			コンテナ番号<br>
			<%=sCntNo%><br>

			完了した作業<br>
			<select name="operation">
				<option value="C">出:空倉庫着</option>
				<option value="D">出:ﾊﾞﾝﾆﾝｸﾞ完</option>
				<option value="A">入:実入倉庫着</option>
				<option value="B">入:デバン完</option>
			</select>
			<br><br>
			<input type="hidden" name="ContNo" value="<%=sCntNo%>">
			<input type="hidden" name="UserID" value="<%=sUserID%>">
			<input type="submit" value="決定">
		</form>
<%
	End If
%>
	<hr>
	</body>
	</html>
<%
End If
%>
