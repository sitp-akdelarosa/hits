<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sUserID
Dim conn, rs
Dim sContE, sContN
Dim sLastContNo
Dim nlen
Dim sSQL
Dim sErrMsg

sErrMsg = ""
sContE = ""
sContN = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

sUserID = Trim(Request.QueryString("UserID"))

sErrMsg = CheckUserID(sUserID)

If sErrMsg = "" Then
	ConnectSvr conn, rs

	' ユーザテーブルを検索し、直前に操作したコンテナ番号を取得する
	sSQL = "SELECT lUserTable.BeforeCntnrNo FROM lUserTable WHERE lUserTable.UserID='" & sUserID & "'"
	rs.Open sSQL, conn, 0, 1
	If rs.Eof Then
		rs.Close
		rs.Open "lUserTable", conn, 2, 2
		rs.AddNew
		rs("UserID") = sUserID
		rs("CompanyName") = "Unknown"
		rs.Update
		rs.Close
	Else
		If Not IsNull(rs("BeforeCntnrNo")) Then
			' コンテナ番号を英字部分と数字部分に分割する
			sLastContNo = rs("BeforeCntnrNo")
			sContE = "value=""" & Left(sLastContNo, 4) & """ "
			nlen = Len(sLastContNo)
			If 4 < nlen Then
				sContN = "value=""" & Right(sLastContNo, nlen - 4) & """ "
			End If
		End If
		rs.Close
	End If
	conn.Close
End If

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If sErrMsg="" Then
	WriteLogM oFS, sUserID, "6200", "携帯-ログイン", "10",sPhoneType, sUserID & "," & "入力内容の正誤:0(正しい)"
	WriteLogM oFS, sUserID, "6201", "携帯-完了時刻コンテナ番号入力", "00",sPhoneType, ","
Else
	WriteLogM oFS, sUserID, "6200", "携帯-ログイン", "10",sPhoneType, sUserID & "," & "入力内容の正誤:1(誤り)" & sErrMsg
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
<%
	If sErrMsg <> "" Then
%>
		<display title="完了時刻入力">
			<center>
			【完了時刻入力】<br><br>
			<center>
			<%=sErrMsg%><br>
			<center>
			<a task="gosub" dest="index.asp">ﾒﾆｭｰ</a>
		</display>
<%
	Else
%>
		<entry name="p1" key="cont_e" format="*A" title="完了時刻入力">
			<action type="accept" task="go" dest="#p2">
			<center>
			【完了時刻入力】<br>
			ｺﾝﾃﾅ番号<br>
			先頭英字4桁:
		</entry>

		<entry name="p2" key="cont_s" format="*N">
			<action type="accept" task="go" dest="mrung02.asp?UserID=<%=sUserID%>&cont_e=$cont_e&cont_s=$cont_s">
			<center>
			【完了時刻入力】<br>
			ｺﾝﾃﾅ番号<br>
			数字部分7桁:
		</entry>
<%
	End If
%>
	</hdml>
<%
Else
	' EzWeb以外のタグを編集
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("完了時刻入力")%>
	</head>
	<body>
	<center>
	【完了時刻入力】
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<br>
		<%=sErrMsg%><br>
		<br>
		<form action="index.asp" method="get">
			<input type="submit" value="ﾒﾆｭｰ">
		</form>
<%
	Else
%>
		<form action="mrung02.asp" method="get">
			ｺﾝﾃﾅ番号入力<br>
			<table boreder="0">
				<tr><td>
					英字4桁:
					<input type="text" name="cont_e" <%=sContE%> maxlength="4" <%=GetTextSizeMode(4, "A")%>><br>
				</td></tr>
				<tr><td>
					数字:
					<input type="text" name="cont_s" <%=sContN%> maxlength="8" <%=GetTextSizeMode(8, "N")%>><br>
				</td></tr>
			</table>
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
