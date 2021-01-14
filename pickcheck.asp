<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="pickcom.inc"-->

<!--#include file="vessel.inc"-->

<%

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '入力画面を記憶
    Session.Contents("findcsv")="no"    ' 直接入力であることを記憶

    Session.Contents("sortkey")="指定日"

    ' 指定引数の取得
    Dim strShipper       '荷主コード
    Dim strTrucker       '陸運コード
    Dim strForwader      '海貨コード
    Dim strVslCode       '船名コード
    Dim strOpeCode       '港運コード
    Dim strVoyCtrl       'Voyage No.
    Dim strPickDate      '空コン搬出指定日
    strShipper = UCase(Trim(Request.QueryString("ninushi")))
    strTrucker = UCase(Trim(Request.QueryString("rikuun")))
    strForwader = UCase(Trim(Request.QueryString("kaika")))
    strVslCode = UCase(Trim(Request.QueryString("vessel")))
    strVoyCtrl = UCase(Trim(Request.QueryString("voyage")))
    strPickDate = Trim(Request.QueryString("decyear")) & "/" & Trim(Request.QueryString("decmon")) & "/" &_
				  Trim(Request.QueryString("decday"))

	Dim iNum,strOption
    strOption = ""
   ' ログイン種別の取得とその処理
    strUserKind=Session.Contents("userkind")
    If strUserKind="海貨" Then
		iNum = "a101"
        strForwader=Session.Contents("userid")
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strShipper & "/" & strTrucker
    ElseIf strUserKind="陸運" Then
		iNum = "a102"
        strTrucker=Session.Contents("userid")
        strOption = strForwader
    ElseIf strUserKind="荷主" Then
		iNum = "a103"
        strShipper=Session.Contents("userid")
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strForwader
    ElseIf strUserKind="港運" Then
		iNum = "a104"
        strOpeCode=Session.Contents("userid")
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strForwader & "/" & strPickDate
    End If

    ' 参照Keyを記憶
    Session.Contents("findkey1")=strShipper       '荷主コード
    Session.Contents("findkey2")=strForwader      '海貨コード
    Session.Contents("findkey3")=strTrucker       '陸運コード
    Session.Contents("findkey4")=strVslCode       '船名コード
    Session.Contents("findkey5")=strVoyCtrl       'Voyage No.
    Session.Contents("findkey6")=strPickDate      '空コン搬出指定日
    Session.Contents("findkey7")=strOpeCode       '港運コード

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' コンテナ情報レコードの取得
    ConnectSvr conn, rsd

    ' 検索条件の作成
    sWhere = ""

    '荷主コード
    If strShipper<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Shipper='" & strShipper & "'"
        Else
            sWhere = "ExportCargoInfo.Shipper='" & strShipper & "'"
        End If
    End If
    '海貨コード
    If strForwader<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Forwarder='" & strForwader & "'"
        Else
            sWhere = "ExportCargoInfo.Forwarder='" & strForwader & "'"
        End If
    End If
    '陸運コード
    If strTrucker<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Trucker='" & strTrucker & "'"
        Else
            sWhere = "ExportCargoInfo.Trucker='" & strTrucker & "'"
        End If
    End If
    '船名コード
    If strVslCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.VslCode='" & strVslCode & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.VslCode='" & strVslCode & "'"
        End If
    End If
    '港運コード
    If strOpeCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.OpeCode='" & strOpeCode & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.OpeCode='" & strOpeCode & "'"
        End If
    End If
    'Voyage No.
    If strVoyCtrl<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.LdVoyage='" & strVoyCtrl & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.LdVoyage='" & strVoyCtrl & "'"
        End If
    End If
   '空コン搬出指定日
    If strPickDate<>"//" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.PickDate<='" & strPickDate & " 23:59:59'" &_
							  " And ExportCargoInfo.PickDate>='" & strPickDate & " 00:00:00'"
        Else
            sWhere = sWhere & "ExportCargoInfo.PickDate<='" & strPickDate & " 23:59:59'" &_
						 " And ExportCargoInfo.PickDate>='" & strPickDate & " 00:00:00'"
        End If
    End If

    sSort="ExportCargoInfo.PickDate"

    ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
    strFileName="./temp/" & strFileName
    ' テンポラリファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSExpCntnr(conn, rsd, ti, sWhere, sSort)

    ti.Close
    conn.Close

    ' Tempファイル属性設定
    SetTempFile "MSEXPORT"

    If bWriteFile = 0 Then
        ' 該当レコードないとき
        bError = true
        strError = "指定条件に該当するコンテナはありませんでした。"
        strOption = strOption & "," & "入力内容の正誤:1(誤り)"
    Else
        strOption = strOption & "," & "入力内容の正誤:0(正しい)"

        ' DT02トランザクションを発行する
        If strUserKind="陸運" Then
            ' トランザクションファイルの拡張子 
            Const SEND_EXTENT = "snd"
            sTranID = "DT02"
            ' 処理区分
            Const sSyori = "R"
            ' 送信場所
            Const sPlace = ""

            ' テンポラリファイル名を作成して、セッション変数に設定
            Dim sDT02, iSeqNo, strFileName_01, sTusin
            'シーケンス番号
            iSeqNo = GetDailyTransNo
            '通信日時取得
            sTusin  = SetTusinDate

            sDT02 = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
                    strTrucker & "," & sPlace & ",X," & strTrucker & "," & strForwader
            sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
            strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
            Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
            ti.WriteLine sDT02
            ti.Close
        End If
    End If

    ' 輸出コンテナ照会
    WriteLog fs, iNum,"空コンピックアップシステム-" & strUserKind & "用照会","10", strOption

    If bError Then
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
<!-------------ここから照会エラー画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<td rowspan=2><%
    If strUserKind="海貨" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf strUserKind="陸運" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf strUserKind="荷主" Then
        Response.Write "<img src='gif/picknit.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/pickkot.gif' width='506' height='73'>"
    End If
%></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
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
		<BR><BR><BR>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>空コンピックアップ情報照会
<%
    If strUserKind="海貨" Then
        Response.Write "(海貨用)"
    ElseIf strUserKind="陸運" Then
        Response.Write "(陸運用)"
    ElseIf strUserKind="荷主" Then
        Response.Write "(荷主用)"
    Else
        Response.Write "(港運用)"
    End If
%>
            </b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>

      <table>
        <tr>
          <td>
<%
    ' エラーメッセージの表示
    DispErrorMessage strError
%>
          </td></tr>
      </table>
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
<!-------------照会エラー画面終わり--------------------------->
<%
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

<%
    Else
        ' 一覧画面へリダイレクト
        If strUserKind="海貨" Then
            Response.Redirect "picklist.asp?kind=1"
        ElseIf strUserKind="陸運" Then
            Response.Redirect "picklist.asp?kind=2"
        ElseIf strUserKind="荷主" Then
            Response.Redirect "picklist.asp?kind=3"
        ElseIf strUserKind="港運" Then
            Response.Redirect "picklist.asp?kind=4"
        End If
    End If
%>
