<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB012.inc"-->
<html>

<head>
<title>搬出入予約申請結果画面</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
/* 戻るのクリック */
function ClickBack(go) {
		location.href = "SYWB010.asp?YMD=" + YMD.value + "&NAME=" + HHName.value + "&Term_Name=" + Term_Name.value
					  + "&Type=" + TRIM(Request.Form("RDType" & 0))
}

</SCRIPT>
</head>

<body >
<%
	Dim sYMD, sHH, sHHName, sTerm_Name, sTerm_CD
	Dim sContNoRec(3), sBKNo(3), sContSizeRec(3), bChkA(3), bChkB(3), bChkC(3), _
		sContNoDel(3), sChID(3), sBLNo(3), sContSizeDel(3), sDeliverTo(3), sRDType(3), sReceiveFrom(3)
'2003/08/27 認証ＩＤ追加(ICCT対応)
	Dim sErrMsg(4), sOpeNoRec(3), sOpeNoDel(3), sErrFlag(3), sNinID(3)
	Dim i, sWk
	Dim conn, rsd
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sRD

	'指定日付取得
	sYMD = TRIM(Request.QueryString("TDATE"))
	sHH  = TRIM(Request.QueryString("HH"))
	sHHName = TRIM(Request.QueryString("HHNAME"))
	sTerm_Name = Trim(Request.QueryString("Term_Name"))		'VP対応
	sTerm_CD = Trim(Request.QueryString("Terminal"))		'VP対応

	'指定値の取得
	For i = 0 To 3
		sWk = CStr(i + 1)
		sRDType(i)      =       TRIM(Request.Form("RDType"      & sWk))
		sContNoRec(i)   = UCASE(TRIM(Request.Form("ContNoRec"   & sWk)))
		sBKNo(i)        = UCASE(TRIM(Request.Form("BKNo"        & sWk)))
		sContSizeRec(i) = UCASE(TRIM(Request.Form("ContSizeRec" & sWk)))
		bChkA(i)        =            Request.Form("checkA"      & sWk) = "on"
		bChkB(i)        =            Request.Form("checkB"      & sWk) = "on"
		bChkC(i)        =            Request.Form("checkC"      & sWk) = "on"
		sContNoDel(i)   = UCASE(TRIM(Request.Form("ContNoDel"   & sWk)))
		sChID(i)        = UCASE(TRIM(Request.Form("ChID"        & sWk)))
		sBLNo(i)        = UCASE(TRIM(Request.Form("BLNo"        & sWk)))
		sContSizeDel(i) =       TRIM(Request.Form("ContSizeDel" & sWk))
		sDeliverTo(i)   = 	TRIM(Request.Form("DeliverTo"   & sWk))
		sReceiveFrom(i) = 	TRIM(Request.Form("ReceiveFrom"   & sWk))
		sNinID(i)       = 	TRIM(Request.Form("NinID"   & sWk))	'2003/08/27
	Next
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

	'予約更新処理	(2003/08/27 認証IDの追加)
	Call UpdateApp(conn, rsd, _
			sUsrID, sGrpID, _
			sYMD, sHH, _
			sRDType, _
			sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC,  _
			sContNoDel, sChID, sBLNo, sContSizeDel, sDeliverTo, sReceiveFrom,   _
			sTerm_CD, sNinID, sErrMsg, sOpeNoRec, sOpeNoDel) 

	'ＤＢ切断
	conn.Close
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title23.gif" width="236" height="34"><p>
<table border="1">   
	<tr ALIGN=middle>
		<td width="120" bgcolor ="#e8ffe8">作業時間</td>
		<td width="360" ><%=ChgYMDStr2(sYMD)%>　<%=sHHName%></td>
	</tr>
	<tr ALIGN=middle>
		<td width="120" bgcolor ="#e8ffe8">搬出入先</td>
		<td width="360" bgcolor ="#ffffff"><%=sTerm_Name%></td>
	</tr>
</table>
</center>
<br>   
	<font face="ＭＳ ゴシック">
<%
	'戻るときのエラー表示のフラグクリア
	For i = 0 to 3
		sErrFlag(i) = " "
	Next 

	If sErrMsg(0) <> "" Then	'登録エラーの場合
		Response.Write "<center>" & sErrMsg(0) & "</center>"
		For i = 0 to 3
			sErrFlag(i) = "E"
		Next 
	Else
		For i = 0 To 3
			If sRDType(i) <> "" Then
				If     sRDType(i) = "DUAL" Then
					sRD = "DUAL"
				ElseIf sRDType(i) = "REC" Then
					sRD = "搬入"
				Else
					sRD = "搬出"
				End If
%>
	<center>
		<table border="0" width="500"  >   
			<tr><th align=left><FONT COLOR="#00008B">＜予約<%=CStr(i + 1)%>＞</FONT></th></tr>
		</table>
		<table border="1" width="500"  >   
			<tr><th width="60" bgcolor ="#40E0D0">種類</th>
			    <td COLSPAN=2 bgcolor="#ffffcc"><%=sRD%></td>
			</tr>
<%
				If sRDType(i) = "DUAL" or sRDType(i) = "REC" Then
					sWk = ""
					If bChkA(i) Then
						sWk = sWk & "搬出を載せない　"
					End If
					If bChkB(i) Then
						sWk = sWk & "夕積のみ載せる　"
					End If
					If bChkC(i) Then
						sWk = sWk & "20/40兼用"
					End If
					If sWk = "" Then
						sWk = "　"
					End If
%>
					<tr><th width="60" bgcolor ="#40E0D0" ROWSPAN=5>搬入</th>
					    <td width="160" bgcolor="#cccc99">コンテナ番号</td>
					    <td><%=sContNoRec(i)%></td>
					</tr>
					<tr><td width="160" bgcolor="#cccc99">ブッキング</td>
					    <td><%=sBKNo(i)%></td>
					</tr>
					<tr><td width="160" bgcolor="#cccc99">サイズ</td>
					    <td><%=sContSizeRec(i)%></td>
					</tr>
					<tr><td width="160" bgcolor="#ffffcc">シャーシ</td>
					    <td><%=sWk%></td>
					</tr>
					<tr><td width="160" bgcolor="#ffffcc">搬入元</td>
<%						If trim(sReceiveFrom(i)) = "" Then %>				
							<td>　</td>
<%						Else %>		
							<td><%=sReceiveFrom(i)%></td>
<%						End If %>		
					</tr>
<%
				End If
				If sRDType(i) = "DUAL" or sRDType(i) = "DEL" Then
%>
					<tr><th width="60" bgcolor ="#40E0D0" ROWSPAN=4>搬出</th>
<%
						If sBLNo(i) = "" Then
%>
							<td width="160" bgcolor ="#cccc99">コンテナ番号</td>
							<td><%=sContNoDel(i)%></td>
<%
						Else
%>
							<td width="160" bgcolor ="#cccc99">ＢＬ番号</td>
							<td><%=sBLNo(i)%></td>
<%
						End If
%>
					</tr>
					<tr><td width="160" bgcolor="#ffffcc">搬出先</td>
<%						If trim(sDeliverTo(i)) = "" Then %>				
							<td>　</td>
<%						Else %>		
							<td><%=sDeliverTo(i)%></td>
<%						End If %>		
					</tr>
					<tr><td width="160" bgcolor="#ffffcc">シャーシID</td>
<%						If trim(sChID(i)) = "" Then %>				
							<td>　</td>
<%						Else %>		
							<td><%=sChID(i)%></td>
<%						End If %>		
					</tr>

					<tr><td width="160" bgcolor="#ffffcc">認証ID</td>
<%						If trim(sNinID(i)) = "" Then %>				
							<td>　</td>
<%						Else %>		
							<td><%=sNinID(i)%></td>
<%						End If %>		
					</tr>
<%
				End If
%>
		</table>
		</center>
		<br>
<%
				If sErrMsg(i + 1) <> "" Then	'登録エラーの場合
					Response.Write "<center><FONT color=Red><U>" & "（結果）：不可　" & sErrMsg(i + 1) & "</U></FONT></center>"
					sErrFlag(i) = "E"
					%><br><% 
				Else							'登録OKの場合
%>
		<center><FONT color=Red><U>（結果）：ＯＫ　</FONT></center><br></U>
		<center>
		<table border="1" width="500"  >   
			<tr bgcolor ="#40E0D0">
						<td align=center>予約番号</td>
					    <td align=center>コンテナ／ＢＬ番号</td>
					    <td align=center>種類</td>
					    <td align=center>サイズ</td>
					    <td align=center>シャーシID</td>
			</tr>
<%
					If sRDType(i) = "DUAL" or sRDType(i) = "REC" Then
%>
			<tr>
						<td><%=sOpeNoRec(i)%></td>
					    <td><%=sContNoRec(i)%></td>
					    <td align=center>搬入</td>
					    <td><%=sContSizeRec(i)%></td>
<%						If trim(sChID(i)) = "" Then %>				
							<td>　</td>
<%						Else %>		
							<td><%=sChID(i)%></td>
<%						End If %>		
			</tr>
<%
					End If
					If sRDType(i) = "DUAL" or sRDType(i) = "DEL" Then
%>
			<tr>
						<td><%=sOpeNoDel(i)%></td>
					    <td><%=sContNoDel(i) & sBLNo(i)%></td>
					    <td align=center>搬出</td>
					    <td><%=sContSizeDel(i)%></td>
<%						If trim(sChID(i)) = "" Then %>				
							<td>　</td>
<%						Else %>		
							<td><%=sChID(i)%></td>
<%						End If %>		
			</tr>
<%
					End If
%>
		</table>
		</center>
<%
				End If
%>
		<br>     
		<br>
<%
			End If
		Next
	End If

	dim sText
'''	sText = "SYWB010.asp?YMD=" & sYMD & sHH & "&NAME=" & sHHName		'VP対応
	sText = "SYWB010.asp?YMD=" & sYMD & sHH & "&NAME=" & sHHName & "&Term_Name=" & sTerm_Name
	sText = sText  & "&Terminal=" & sTerm_CD

	if	sErrFlag(0) = "E" Then
		sText = sText & "&sRDType1=" &  sRDType(0) 
		sText = sText & "&sContNoRec1=" &  sContNoRec(0) 
		sText = sText & "&sBKNo1=" &  sBKNo(0) 
		sText = sText & "&sContSizeRec1=" &  sContSizeRec(0) 
		sText = sText & "&bChkA1=" &  bChkA(0) 
		sText = sText & "&bChkB1=" &  bChkB(0) 
		sText = sText & "&bChkC1=" &  bChkC(0) 
		sText = sText & "&sContNoDel1=" &  sContNoDel(0) 
		sText = sText & "&sChID1=" &  sChID(0) 
		sText = sText & "&sBLNo1=" &  sBLNo(0) 
		sText = sText & "&sContSizeDel1=" &  sContSizeDel(0) 
		sText = sText & "&sDeliverTo1=" &  sDeliverTo(0)
		sText = sText & "&sReceiveFrom1=" &  sReceiveFrom(0)				
		sText = sText & "&sNinID1=" &  sNinID(0)				
	Else
		sText = sText & "&sRDType1=" &  ""
		sText = sText & "&sContNoRec1=" &  ""
		sText = sText & "&sBKNo1=" &  ""
		sText = sText & "&sContSizeRec1=" &  ""
		sText = sText & "&bChkA1=" &  ""
		sText = sText & "&bChkB1=" &  ""
		sText = sText & "&bChkC1=" &  ""
		sText = sText & "&sContNoDel1=" &  ""
		sText = sText & "&sChID1=" &  ""
		sText = sText & "&sBLNo1=" &  ""
		sText = sText & "&sContSizeDel1=" &  ""
		sText = sText & "&sDeliverTo1=" &  ""
		sText = sText & "&sReceiveFrom1=" &  ""
		sText = sText & "&sNinID1=" &  ""
	End IF

	if	sErrFlag(1) = "E" Then
		sText = sText & "&sRDType2=" &  sRDType(1) 
		sText = sText & "&sContNoRec2=" &  sContNoRec(1) 
		sText = sText & "&sBKNo2=" &  sBKNo(1) 
		sText = sText & "&sContSizeRec2=" &  sContSizeRec(1) 
		sText = sText & "&bChkA2=" &  bChkA(1) 
		sText = sText & "&bChkB2=" &  bChkB(1) 
		sText = sText & "&bChkC2=" &  bChkC(1) 
		sText = sText & "&sContNoDel2=" &  sContNoDel(1) 
		sText = sText & "&sChID2=" &  sChID(1) 
		sText = sText & "&sBLNo2=" &  sBLNo(1) 
		sText = sText & "&sContSizeDel2=" &  sContSizeDel(1) 
		sText = sText & "&sDeliverTo2=" &  sDeliverTo(1) 
		sText = sText & "&sReceiveFrom2=" &  sReceiveFrom(1)				
		sText = sText & "&sNinID2=" &  sNinID(1)				
	Else
		sText = sText & "&sRDType2=" &  ""
		sText = sText & "&sContNoRec2=" &  ""
		sText = sText & "&sBKNo2=" &  ""
		sText = sText & "&sContSizeRec2=" &  ""
		sText = sText & "&bChkA2=" &  ""
		sText = sText & "&bChkB2=" &  ""
		sText = sText & "&bChkC2=" &  ""
		sText = sText & "&sContNoDel2=" &  ""
		sText = sText & "&sChID2=" &  ""
		sText = sText & "&sBLNo2=" &  ""
		sText = sText & "&sContSizeDel2=" &  ""
		sText = sText & "&sDeliverTo2=" &  ""
		sText = sText & "&sReceiveFrom2=" &  ""
		sText = sText & "&sNinID2=" &  ""
	End IF

	if	sErrFlag(2) = "E" Then	
		sText = sText & "&sRDType3=" &  sRDType(2) 
		sText = sText & "&sContNoRec3=" &  sContNoRec(2) 
		sText = sText & "&sBKNo3=" &  sBKNo(2) 
		sText = sText & "&sContSizeRec3=" &  sContSizeRec(2) 
		sText = sText & "&bChkA3=" &  bChkA(2) 
		sText = sText & "&bChkB3=" &  bChkB(2) 
		sText = sText & "&bChkC3=" &  bChkC(2) 
		sText = sText & "&sContNoDel3=" &  sContNoDel(2) 
		sText = sText & "&sChID3=" &  sChID(2) 
		sText = sText & "&sBLNo3=" &  sBLNo(2) 
		sText = sText & "&sContSizeDel3=" &  sContSizeDel(2) 
		sText = sText & "&sDeliverTo3=" &  sDeliverTo(2) 
		sText = sText & "&sReceiveFrom3=" &  sReceiveFrom(2)				
		sText = sText & "&sNinID3=" &  sNinID(2)				
	Else
		sText = sText & "&sRDType3=" &  ""
		sText = sText & "&sContNoRec3=" &  ""
		sText = sText & "&sBKNo3=" &  ""
		sText = sText & "&sContSizeRec3=" &  ""
		sText = sText & "&bChkA3=" &  ""
		sText = sText & "&bChkB3=" &  ""
		sText = sText & "&bChkC3=" &  ""
		sText = sText & "&sContNoDel3=" &  ""
		sText = sText & "&sChID3=" &  ""
		sText = sText & "&sBLNo3=" &  ""
		sText = sText & "&sContSizeDel3=" &  ""
		sText = sText & "&sDeliverTo3=" &  ""
		sText = sText & "&sReceiveFrom3=" &  ""
		sText = sText & "&sNinID3=" &  ""
	End IF

	if	sErrFlag(3) = "E" Then	
		sText = sText & "&sRDType4=" &  sRDType(3) 
		sText = sText & "&sContNoRec4=" &  sContNoRec(3) 
		sText = sText & "&sBKNo4=" &  sBKNo(3) 
		sText = sText & "&sContSizeRec4=" &  sContSizeRec(3) 
		sText = sText & "&bChkA4=" &  bChkA(3) 
		sText = sText & "&bChkB4=" &  bChkB(3) 
		sText = sText & "&bChkC4=" &  bChkC(3) 
		sText = sText & "&sContNoDel4=" &  sContNoDel(3) 
		sText = sText & "&sChID4=" &  sChID(3) 
		sText = sText & "&sBLNo4=" &  sBLNo(3) 
		sText = sText & "&sContSizeDel4=" &  sContSizeDel(3) 
		sText = sText & "&sDeliverTo4=" &  sDeliverTo(3) 
		sText = sText & "&sReceiveFrom4=" &  sReceiveFrom(3)
		sText = sText & "&sNinID4=" &  sNinID(3)				
	Else
		sText = sText & "&sRDType4=" &  ""
		sText = sText & "&sContNoRec4=" &  ""
		sText = sText & "&sBKNo4=" &  ""
		sText = sText & "&sContSizeRec4=" &  ""
		sText = sText & "&bChkA4=" &  ""
		sText = sText & "&bChkB4=" &  ""
		sText = sText & "&bChkC4=" &  ""
		sText = sText & "&sContNoDel4=" &  ""
		sText = sText & "&sChID4=" &  ""
		sText = sText & "&sBLNo4=" &  ""
		sText = sText & "&sContSizeDel4=" &  ""
		sText = sText & "&sDeliverTo4=" &  ""
		sText = sText & "&sReceiveFrom4=" &  ""
		sText = sText & "&sNinID4=" &  ""
	End IF

	
%>
<br>     
<center>
<table border=0>
    <form  METHOD="post"  NAME="Sub" ACTION="<%=sText%>" >
	<td><input type="submit" value="　戻る　" id=submit1 name=submit1></td>
	</form>

    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
	<td><input type="submit" value="一覧画面へ" id=submit2 name=submit2></td>
	</form>
</table>
</center>

</body>     
</html>     

