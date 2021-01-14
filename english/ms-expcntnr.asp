<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ms-ExpCom.inc"-->

<!--#include file="vessel.inc"-->

<%
    ' セッションのチェック
    CheckLogin "expentry.asp"

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '入力画面を記憶
    Session.Contents("findcsv")="no"    ' 直接入力であることを記憶

    ' 指定引数の取得
    Dim strShipper       '荷主コード
    Dim strTrucker       '陸運コード
    Dim strForwader      '海貨コード
    Dim strVslCode       '船名コード
    Dim strVoyCtrl       'Voyage No.
    strShipper = UCase(Trim(Request.QueryString("ninushi")))
    strTrucker = UCase(Trim(Request.QueryString("rikuun")))
    strForwader = UCase(Trim(Request.QueryString("kaika")))
    strVslCode = UCase(Trim(Request.QueryString("vessel")))
    strVoyCtrl = UCase(Trim(Request.QueryString("voyage")))

	Dim iNum,strOption
    strOption = ""
   ' ログイン種別の取得とその処理
    strUserKind=Session.Contents("userkind")
    If strUserKind="海貨" Then
		iNum = "1101"
        strForwader=Session.Contents("userid")
        Session.Contents("sortkey")="荷主名"           ' ソートキーを指定
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strShipper & "/" & strTrucker
    ElseIf strUserKind="陸運" Then
		iNum = "1102"
        strTrucker=Session.Contents("userid")
        Session.Contents("sortkey")="海貨"             ' ソートキーを指定
        strOption = strForwader
    ElseIf strUserKind="荷主" Then
		iNum = "1103"
        strShipper=Session.Contents("userid")
        Session.Contents("sortkey")="荷主管理番号"     ' ソートキーを指定
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strForwader
    End If

    ' 参照Keyを記憶
    Session.Contents("findkey1")=strShipper       '荷主コード
    Session.Contents("findkey2")=strForwader      '海貨コード
    Session.Contents("findkey3")=strTrucker       '陸運コード
    Session.Contents("findkey4")=strVslCode       '船名コード
    Session.Contents("findkey5")=strVoyCtrl       'Voyage No.

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
    'Voyage No.
    If strVoyCtrl<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.LdVoyage='" & strVoyCtrl & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.LdVoyage='" & strVoyCtrl & "'"
        End If
    End If

    ' Sort条件の作成
    strSortKey=Session.Contents("sortkey")
    If strSortKey="荷主名" Then
        sSort="ExportCargoInfo.Shipper, ExportCargoInfo.ShipCtrl"
    ElseIf strSortKey="海貨" Then
        sSort="ExportCargoInfo.Forwarder"
    ElseIf strSortKey="荷主管理番号" Then
        sSort="ExportCargoInfo.ShipCtrl"
    ElseIf strSortKey="倉庫到着" Then
        sSort="ExportCargoInfo.WHArTime"
    ElseIf strSortKey="CY到着" Then
        sSort="ExportCargoInfo.CYRecDate"
    ElseIf strSortKey="陸運業者" Then
'        sSort="mTrucker.FullName"
        sSort="ExportCargoInfo.Trucker"
    End If

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
    WriteLog fs, iNum,"輸出コンテナ照会-" & strUserKind & "用照会","10", strOption

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
        Response.Write "<img src='gif/expkaika.gif' width='506' height='73'>"
    ElseIf strUserKind="陸運" Then
        Response.Write "<img src='gif/exprikuun.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/expninushi.gif' width='506' height='73'>"
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
          <td nowrap><b>輸出コンテナ情報照会
<%
    If strUserKind="海貨" Then
        Response.Write "(海貨用)"
    ElseIf strUserKind="陸運" Then
        Response.Write "(陸運用)"
    Else
        Response.Write "(荷主用)"
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
            Response.Redirect "ms-explist1.asp"          '輸出コンテナ一覧
        ElseIf strUserKind="陸運" Then
            Response.Redirect "ms-explist2.asp"          '輸出コンテナ一覧
        ElseIf strUserKind="荷主" Then
            Response.Redirect "ms-explist3.asp"          '輸出コンテナ一覧
        End If
    End If
%>
