<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ms-ExpCom.inc"-->

<%
    ' ログイン種別の取得とその処理
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"         'メニュー
        Response.End
    End If

    ' セッションのチェック
    If strUserKind="海貨" Then
        CheckLogin "ms-expentry.asp?kind=1"
    ElseIf strUserKind="陸運" Then
        CheckLogin "ms-expentry.asp?kind=2"
    Else
        CheckLogin "ms-expentry.asp?kind=3"
    End If

    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' 記憶している検索条件をロード
    strShipper=Session.Contents("findkey1")       '荷主コード
    strForwader=Session.Contents("findkey2")      '海貨コード
    strTrucker=Session.Contents("findkey3")       '陸運コード
    strVslCode=Session.Contents("findkey4")       '船名コード
    strVoyCtrl=Session.Contents("findkey5")       'Voyage No.

    ' 指定引数の取得
    Dim strRequest
    strRequest = Request.QueryString("request")  ' 更新リクエスト画面ID
    Dim strSortKey
    strSortKey = Request.QueryString("sort")     ' ソートモードの取得
    If strSortKey="" Then
        strSortKey=Session.Contents("sortkey")
    End If
    Session.Contents("sortkey")=strSortKey

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
    sSort = ""

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
    ' 転送ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSExpCntnr(conn, rsd, ti, sWhere, sSort)

    ' ファイルとDBのクローズ
    ti.Close
    conn.Close

    ' 詳細画面からのとき、該当コンテナデータの行を検索する
    If strRequest="ms-expdetail.asp" Then
        ' 記憶している検索条件をロード
        strFindCntnr=Session.Contents("dispexpctrl")     ' 表示荷主管理番号

        ' 表示ファイルのOpen
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

        ' 詳細表示行のデータの取得
        bWriteFile = 0                    '検索結果フラグ
        LineNo=0
        Do While Not ti.AtEndOfStream
            anyTmp=Split(ti.ReadLine,",")
            LineNo=LineNo+1
            If anyTmp(14)=strFindCntnr Then
               bWriteFile=1
               Exit Do
            End If
        Loop

        ti.Close
    End If

    If bWriteFile = 0 Then
        ' 該当レコードないとき
        bError = true
        strError = "指定条件に該当するコンテナはなくなりました。"
    End If

    ' 輸出コンテナ照会
'    WriteLog fs, "輸出入業務支援-輸出コンテナ照会", "画面更新:SortKey," & strSortKey

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
        Response.Write "<img src='gif/expkaika.gif' width='506' height='73'>"
    ElseIf strUserKind="陸運" Then
        Response.Write "<img src='gif/exprikuun.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/expninushi.gif' width='506' height='73'>"
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
          <td nowrap>
            <dl> 
            <dt><font color="#000066" size="+1">【輸出コンテナ情報
<%
    If strUserKind="海貨" Then
        Response.Write "(海貨用)"
    ElseIf strUserKind="陸運" Then
        Response.Write "(陸運用)"
    Else
        Response.Write "(荷主用)"
    End If
%>
               画面】</font><br>
            <dd>
<%
    ' エラーメッセージの表示
    DispErrorMessage strError
%>
            </dl>
          </td>
        </tr>
      </table>
      <form action="ms-expentry.asp">
        <br><br>
        <input type="submit" value=" 戻  る ">
      </form>
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
    DispMenuBarBack "ms-expentry.asp"
%>
</body>
</html>

<%
    Else
        If strRequest="ms-expdetail.asp" Then
            ' 詳細画面へリダイレクト
            Response.Redirect "ms-expdetail.asp?line=" & LineNo  '輸出コンテナ詳細
        Else
            ' 一覧画面へリダイレクト
            Response.Redirect strRequest                         '輸出コンテナ一覧
        End If
    End If
%>
