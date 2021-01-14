<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
''    ' セッションのチェック
''    CheckLogin "expentry.asp"

    '入力画面を記憶
    Session.Contents("findcsv")="yes"    ' CSVファイル入力であることを記憶

    ' 指定引数の取得
    Dim strKind
    strKind = Request.QueryString("kind")

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' 参照モードをセッション変数に設定
    If strKind="cntnr" Then
        Session.Contents("findkind")="Cntnr"
    Else
        Session.Contents("findkind")="Booking"
    End If

    ' 転送ファイルの取得
    tb=Request.TotalBytes      :' ブラウザからのトータルサイズ
    br=Request.BinaryRead(tb)  :' ブラウザからの生データ

    ' BASP21 コンポーネントの作成
    Set bsp=Server.CreateObject("basp21")

    filesize=bsp.FormFileSize(br,"csvfile")
    filename=bsp.FormFileName(br,"csvfile")

'    fpath=fs.GetFileName(filename)
    fpath=GetNumStr(Session.SessionID, 8) & "c.csv"
    fpath=fs.BuildPath(Server.MapPath("./temp"),fpath)

    lng=bsp.FormSaveAs(br,"csvfile",fpath)

    ' ファイル転送に失敗したとき
    If lng<=0 Then
        bError=true
        strError = "'" & filename & "'ファイルの転送に失敗しました。"
    Else
        Dim strCntnrNo()

        ' 転送ファイルのOpen
        Set ti=fs.OpenTextFile(fpath,1,True)

        iRecCount=0
        strFindKey=""
        ' 転送ファイルのレコードがある間繰り返す
        Do While Not ti.AtEndOfStream
            cntnrNo = Trim(ti.ReadLine)
            If cntnrNo<>"" Then
                ReDim Preserve strCntnrNo(iRecCount)
                strCntnrNo(iRecCount) = UCase(cntnrNo)
                If strFindKey<>"" Then
                    strFindKey=strFindKey & "," & strCntnrNo(iRecCount)
                Else
                    strFindKey=strCntnrNo(iRecCount)
                End If
                iRecCount=iRecCount + 1
            End If
        Loop
        ti.Close
        Session.Contents("findkey")=strFindKey     ' 参照Keyを記憶
        ' 転送ファイルの削除
        fs.DeleteFile fpath

        ' コンテナ情報レコードの取得
        ConnectSvr conn, rsd

        ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
        strFileName="./temp/" & strFileName
        ' テンポラリファイルのOpen
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
        bWriteFile = 0

        For iCount=0 To iRecCount - 1
            If strKind="cntnr" Then
                sWhere = "ExportCont.ContNo='" & strCntnrNo(iCount) & "'"
            Else
                sWhere = "ExportCont.BookNo='" & strCntnrNo(iCount) & "'"
            End If

            bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
        Next

        ti.Close
        conn.Close

        If bWriteFile = 0 Then
            ' 該当レコードないとき
            bError = true
            strError = "指定条件に該当するコンテナはありませんでした。"
        End If
    End If

    ' Tempファイル属性設定
    SetTempFile "EXPORT"

    strOption = filename

    If bError Then
        strOption = strOption & "," & "入力内容の正誤:1(誤り)"
    Else
        strOption = strOption & "," & "入力内容の正誤:0(正しい)"
    End If

	Dim iWrkNum
    If strKind="cntnr" Then
		iWrkNum = 21
	Else
		iWrkNum = 22
	End If
    ' 輸出コンテナ照会
    WriteLog fs, "1003","輸出コンテナ照会-CSVファイル転送",iWrkNum, strOption

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
          <td rowspan=2><img src="../gif/expentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
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
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>CSVファイル転送</b></td>
          <td><img src="../gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr> 
          <td nowrap>
            <font color="#000066" size="+1">【コンテナ情報照会用ファイル転送画面】</font>
			<BR><br>
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
    DispMenuBarBack "expcsv.asp"
%>
</body>
</html>

<%
    Else
        If bWriteFile = 1 Then
            '戻り画面種別を記憶
            Session.Contents("dispreturn")=0
            ' 詳細画面へリダイレクト
            Response.Redirect "expdetail.asp?line=1"    '輸出コンテナ詳細
        Else
            '戻り画面種別を記憶
            Session.Contents("dispreturn")=0
            ' 一覧画面へリダイレクト
            Response.Redirect "explist.asp"             '輸出コンテナ一覧
        End If
    End If
%>
