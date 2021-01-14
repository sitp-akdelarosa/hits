<!--<%@ LANGUAGE="VBScript" %>-->
<%
'Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="Sywb060.inc"-->
<html>

<head>
<title>空バン予約結果画面</title>
</head>

<body>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
		<p align="center"><img border="0" src="image/title34.gif" width="236" height="34"><p>
<%
	Dim sYMD, Idx
	Dim conn, rsd
	Dim sName, sErrMsg
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sVPBook, sVSL, sVOY, sLINE, sTERM, sSIZE, sTYPE
	Dim sHIGHT, sMATERIAL
	Dim sTERM_Name, sTYPE_Name, sMATERIAL_Name
	Dim sContType_M(14), sTerminal_M(49), sMaterial_M(9)
	Dim sOpeNo
	Dim sDeliverTo

	sOpeNo = "10023"
	'指定日付取得(次画面引継)
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD, 9, 2)
	sYMD = Left(sYMD, 8)
	'作業時間帯(次画面引継)
	sName = TRIM(Request.QueryString("NAME"))
	'選択ブッキング番号
	sVPBook = TRIM(Request.QueryString("BOOK"))
	'選択本船
	sVSL = TRIM(Request.QueryString("VSL"))
	'選択次航
	sVOY = TRIM(Request.QueryString("VOY"))
	'選択船社
	sLINE = TRIM(Request.QueryString("LINE"))
	'選択ターミナル
	sTERM = TRIM(Request.QueryString("TERM"))
	'選択サイズ
	sSIZE = TRIM(Request.QueryString("SIZE"))
	'選択タイプ
	sTYPE = TRIM(Request.QueryString("TYPE"))
	'選択ハイト
	sHIGHT = TRIM(Request.QueryString("HIGHT"))
	'選択材質
	sMATERIAL = TRIM(Request.QueryString("MATERIAL"))
	'コンテナ搬出先
	sDeliverTo = TRIM(Request.QueryString("DELIVERTO"))
%>		<center>
		<table border="1">   
			<tr ALIGN=middle>
				<td width="120" bgcolor ="#e8ffe8">作業時間</td>
				<td width="360" ><%=ChgYMDStr2(sYMD)%>　<%=sName%></td>
			</tr>
		</table>
		<br>
<%
	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	If sGrpID = "" Then
		Response.Write "ユーザが登録されていません。(" & sUsrID & ")"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.End
	End If

	'ターミナル名称の取得
	sTERM_Name =  GetTerminal_Name(conn, rsd, sTERM)
	'コンテナタイプ名称の取得
	sTYPE_Name =  GetContType_Name(conn, rsd, sTYPE)
	'コンテナ材質名称の取得
	sMATERIAL_Name =  GetMaterial_Name(conn, rsd, sMATERIAL)
%>
	<table border="1" width="500"  >
		<tr><td width="120" bgcolor="#cccc99">ブッキング番号</td>
		<td><%=sVPBook%></td>
		</tr>
		<tr>
			<td width="120" bgcolor="#cccc99">対象バンプール</td>
			<td><%=sTERM_Name%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">サイズ</td>
			<td><%=sSIZE%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">タイプ</td>
			<td><%=sTYPE_Name%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">高さ</td>
			<td><%=sHIGHT%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">材質</td>
			<td><%=sMATERIAL_Name%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">本船</td>
			<td><%=sVSL%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">次航</td>
			<td><%=sVOY%></td>
		</tr>
		<tr>
		    <td width="120" bgcolor="#cccc99">船社</td>
			<td><%=sLINE%></td>
		</tr>
	</table><br>
<%

	'空バン予約更新処理
	Call UpdateApp_VPBook(conn, rsd, sUsrID, sGrpID, _
			sYMD, sHH, sVSL, sVOY, sLINE, sVPBook, _
            sTERM, sSIZE, sTYPE, sHIGHT, sMATERIAL, sDeliverTo, _
            sErrMsg, sOpeNoVan) 

'データがない場合ほかのエラー
	if sErrMsg <> "" then
%>		<center>	<%
			Response.Write "<center><FONT color=Red><U>" & "（結果）：不可　" & sErrMsg & "</U></FONT></center>"
			%><br>
		</center>

<%	Else	%>
		<center><FONT  size=4 color=blue><U>（結果）：ＯＫ　予約番号：<%=sOpeNoVan%></FONT></center><br></U>

		<center>
			<table>
			<tr>
			<td><font color=red>（注意）<U>本数不足の理由で予約不可になる可能性があります。１０分後以降に一覧画</U></font><br></td>
			<tr>
			<td>　　　　 <font color=red><u>面を再確認していただきますようお願い致します。</u></font><br></td>
			</tr>
			<tr>
			<td>　　　　<font color=red>（直接引き取りとの重複があった場合）</font><br></td></tr>
			</table>
		</center>

<%
	End If 
	conn.Close
%>
	<br>
	<center>
	<table border=0>
	    <form id=form1 name=form1>
	    <td><input type="button" value="　戻る　" onclick="history.back()"  id=button1 name=button1></td>
		</form>

	    <form  METHOD="post"  NAME="BACK" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="一覧画面へ" id=submit2 name=submit2></td>
		</form>
	</table>
	</center>

</body>
</html>
