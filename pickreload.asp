<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="pickcom.inc"-->

<%
    ' ログイン種別の取得とその処理
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"         'メニュー
        Response.End
    End If

	sSortKey = Request.QueryString("sort")

    ' セッションのチェック
	Dim sLoginKind
    If strUserKind="海貨" Then
        CheckLogin "picklist.asp?kind=1"
		sLoginKind = "1"
    ElseIf strUserKind="陸運" Then
        CheckLogin "picklist.asp?kind=2"
		sLoginKind = "2"
    ElseIf strUserKind="荷主" Then
        CheckLogin "picklist.asp?kind=3"
		sLoginKind = "3"
    Else
        CheckLogin "picklist.asp?kind=4"
		sLoginKind = "4"
    End If

    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' 記憶している検索条件をロード
    strShipper=Session.Contents("findkey1")       '荷主コード
    strForwader=Session.Contents("findkey2")      '海貨コード
    strTrucker=Session.Contents("findkey3")       '陸運コード
    strVslCode=Session.Contents("findkey4")       '船名コード
    strVoyCtrl=Session.Contents("findkey5")       'Voyage No.
    strPickDate=Session.Contents("findkey6")      '空コン搬出日
    strOpeCode=Session.Contents("findkey7")       '港運コード

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"         'メニュー
        Response.End
    End If

    ' データベースの接続
    ConnectSvr conn, rsd

    ' 検索条件の作成
    sWhere = ""

    '荷主コード
    If strShipper<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Shipper='" & strShipper & "'"
            strOption = strOption & ",荷主コード," & strShipper
        Else
            sWhere = "ExportCargoInfo.Shipper='" & strShipper & "'"
            strOption = "荷主コード," & strShipper
        End If
    End If
    '海貨コード
    If strForwader<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Forwarder='" & strForwader & "'"
            strOption = strOption & ",海貨コード," & strForwader
        Else
            sWhere = "ExportCargoInfo.Forwarder='" & strForwader & "'"
            strOption = "海貨コード," & strForwader
        End If
    End If
    '陸運コード
    If strTrucker<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Trucker='" & strTrucker & "'"
            strOption = strOption & ",陸運コード," & strTrucker
        Else
            sWhere = "ExportCargoInfo.Trucker='" & strTrucker & "'"
            strOption = "陸運コード," & strTrucker
        End If
    End If
    '船名コード
    If strVslCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.VslCode='" & strVslCode & "'"
            strOption = strOption & ",船名コード," & strVslCode
        Else
            sWhere = sWhere & "ExportCargoInfo.VslCode='" & strVslCode & "'"
            strOption = "船名コード," & strVslCode
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
            sWhere = sWhere & " And ExportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
            strOption = strOption & ",Voyage No.," & strVoyCtrl
        Else
            sWhere = sWhere & "ExportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
            strOption = "Voyage No.," & strVoyCtrl
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

    ' Sort条件の作成
	If sSortKey="海貨" Then
		sSort="ExportCargoInfo.Forwarder,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="海貨"
	ElseIf sSortKey="荷主" Then
		sSort="ExportCargoInfo.Shipper,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="荷主"
	ElseIf sSortKey="陸運" Then
		sSort="ExportCargoInfo.Trucker,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="陸運"
	ElseIf sSortKey="港運" Then
		sSort="ExportCargoInfo.OpeCode,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="港運"
	Else
		sSort="ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="指定日"
	End If

    ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
    strFileName="./temp/" & strFileName
    ' 転送ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSExpCntnr(conn, rsd, ti, sWhere, sSort)

    ' ファイルとDBのクローズ
    ti.Close
    conn.Close

    If bWriteFile = 0 Then
        ' 該当レコードないとき
        bError = true
        strError = "指定条件に該当するコンテナはなくなりました。"
    End If


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
<!-------------ここからエラー画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2>
<%
    If strUserKind="海貨" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf strUserKind="陸運" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf strUserKind="荷主" Then
        Response.Write "<img src='gif/picknit.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/pickkot.gif' width='506' height='73'>"
    End If
%>
          </td>
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
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>空コンピックアップ情報一覧
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
          <td nowrap>
<%
    ' エラーメッセージの表示
    DispErrorMessage strError
%>
          </td>
        </tr>
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
<!-------------エラー画面終わり--------------------------->
<%
    DispMenuBarBack "pickselect.asp"
%>
</body>
</html>

<%
    Else
         ' 一覧画面へリダイレクト
        Response.Redirect "picklist.asp?kind=" & sLoginKind
    End If
%>
