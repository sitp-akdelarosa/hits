<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="vessel.inc"-->

<%
	' トランザクションファイルの拡張子 
	Const SEND_EXTENT = "snd"
	' トランザクションＩＤ
	Const sTranID = "EX15"
	' 処理区分
	Const sSyori = "R"
	' 送信場所
	Const sPlace = ""
    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"

	sSosin = Trim(Session.Contents("userid"))	'海貨コード
    ' エラーフラグのクリア
    bError = false
    ' 入力フラグのクリア
    bInput = true
    ' 指定引数の取得
    Dim sCntnrNo
    sCntnrNo = UCase(Trim(Request.form("CntnrNo")))
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 入力コンテナNo.のチェック
	ConnectSvr conn, rsd
	sql = "SELECT ExportCont.VslCode, ExportCont.VoyCtrl, VslSchedule.ShipLine, VslSchedule.LdVoyage" & _
	      " FROM ExportCont, VslSchedule" & _
	      " WHERE ExportCont.ContNo = '" & sCntnrNo & "' And VslSchedule.VslCode=ExportCont.VslCode And " & _
          "VslSchedule.VoyCtrl=ExportCont.VoyCtrl"
			 
	'SQLを発行して輸出コンテナを検索
	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
	    sVslCode = Trim(rsd("VslCode"))		'船名
	    sVoyCtrl = Trim(rsd("LdVoyage"))	'次航
	Else
	    ' 該当レコードのないとき エラーメッセージを表示
	    bError = true
	    strError = "該当する輸出コンテナNo.が存在しません。"
	End If
	rsd.Close
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/kaika2t.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
		<table>
			<tr> 
				<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
				<td nowrap><b>コンテナ情報入力</b></td>
				<td><img src="gif/hr.gif"></td>
			</tr>
		</table>
		<BR>
<%

	Dim sLogDate
	sLogDate = Trim(Request.form("Year")) & "年"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Month")),2) & "月"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("Day")),2) & "日"

    If bError Then
	    ' エラーメッセージの表示
	    DispErrorMessage strError 
        strOption = sCntnrNo & "/" & sLogDate & "," & "入力内容の正誤:1(誤り)"

    Else
		'トランザクションファイル作成
	    ' テンポラリファイル名を作成して、セッション変数に設定
	    Dim sEX15, iSeqNo_EX15, strFileName, sTran, sTusinsDate, sDate
		'シーケンス番号
		iSeqNo_EX15 = GetDailyTransNo
		'通信日時取得
		sTusin  = SetTusinDate
		sDate = Trim(Request.form("Year"))
		sDate = sDate & Right("0" & Trim(Request.form("Month")),2)
		sDate = sDate & Right("0" & Trim(Request.form("Day")),2) & "2359"

		sEX15 = iSeqNo_EX15 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sCntnrNo & "," & sDate
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX15
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sEX15
	    ti.Close
		Set ti = Nothing
	    ' エラーメッセージの表示
		strError = "正常に更新されました。"
        If strError="正常に更新されました。" Then
            DispInformationMessage strError
            strOption = sCntnrNo & "/" & sLogDate & "," & "入力内容の正誤:0(正しい)"
        Else
            DispErrorMessage strError
            strOption = sCntnrNo & "/" & sLogDate & "," & "入力内容の正誤:1(誤り)"
        End If

    End If
%>
			</center>
			<br>
		</td>
	</tr>
	<tr>
		<td valign="bottom"> 
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>
<%
    ' 海貨CY搬入時刻指示
    WriteLog fs, "4003","海貨入力CY搬入日","10", strOption
%>
