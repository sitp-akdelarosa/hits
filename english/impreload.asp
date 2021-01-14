<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ImpCom.inc"-->

<%
''    ' セッションのチェック
''    CheckLogin "expentry.asp"

    ' Tempファイル属性のチェック
    CheckTempFile "IMPORT", "impentry.asp"

    ' 記憶している検索条件をロード
    strFindKind=Session.Contents("findkind")     ' 検索条件
    strFindCSV=Session.Contents("findcsv")       ' 検索種別
    strFindKey=Session.Contents("findkey")       ' 検索キー

    ' 指定引数の取得
    Dim strRequest
    strRequest = Request.QueryString("request")  ' 更新リクエスト画面ID

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "expentry.asp"         '輸出コンテナ照会トップ
        Response.End
    End If
    strFileName="../temp/" & strFileName

    ' データベースの接続
    ConnectSvr conn, rsd

    ' 転送ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    ' 検索画面の判別
    If strFindCSV="no" Then
        ' 画面入力のとき
        sWhere = ""
        If strFindKind="Blno" Then               ' Bl番号検索のとき
            iCanma = InStr(strFindKey,",")
            Do While iCanma>0
                sTemp = Trim(Left(strFindKey,iCanma-1))
                strFindKey = Right(strFindKey,Len(strFindKey)-iCanma)
                If sWhere<>"" Then
                    sWhere = sWhere & " Or ImportCont.BLNo='" & sTemp & "'"
                Else
                    sWhere = "ImportCont.BLNo='" & sTemp & "'"
                End If
                iCanma = InStr(strFindKey,",")
            Loop
            If sWhere<>"" Then
                sWhere = sWhere & " Or ImportCont.BLNo='" & Trim(strFindKey) & "'"
            Else
                sWhere = "ImportCont.BLNo='" & Trim(strFindKey) & "'"
            End If
        Else                                     ' Container番号検索のとき
            iCanma = InStr(strFindKey,",")
            Do While iCanma>0
                sTemp = Trim(Left(strFindKey,iCanma-1))
                strFindKey = Right(strFindKey,Len(strFindKey)-iCanma)
                If sWhere<>"" Then
                    sWhere = sWhere & " Or ImportCont.ContNo='" & sTemp & "'"
                Else
                    sWhere = "ImportCont.ContNo='" & sTemp & "'"
                End If
                iCanma = InStr(strFindKey,",")
            Loop
            If sWhere<>"" Then
                sWhere = sWhere & " Or ImportCont.ContNo='" & Trim(strFindKey) & "'"
            Else
                sWhere = "ImportCont.ContNo='" & Trim(strFindKey) & "'"
            End If
        End If

        bWriteFile = SerchImpCntnr(conn, rsd, ti, sWhere)

    Else
        ' 検索キーの分解
        strCntnrNo=Split(strFindKey, ",")
        iRecCount=Ubound(strCntnrNo)+1

        bWriteFile = 0

        For iCount=0 To iRecCount - 1
            If strFindKind="Cntnr" Then
                sWhere = "ImportCont.ContNo='" & strCntnrNo(iCount) & "'"
            Else
                sWhere = "ImportCont.BLNo='" & strCntnrNo(iCount) & "'"
            End If

            bWriteFile = bWriteFile + SerchImpCntnr(conn, rsd, ti, sWhere)
        Next

    End If

    ' ファイルとDBのクローズ
    ti.Close
    conn.Close

    ' 詳細画面からのとき、該当コンテナデータの行を検索する
    If strRequest="impdetail.asp" Then
        ' 記憶している検索条件をロード
        strFindCntnr=Session.Contents("dispcntnr")     ' 表示コンテナNo.

        ' 表示ファイルのOpen
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

        ' 詳細表示行のデータの取得
        bWriteFile = 0                    '検索結果フラグ
        LineNo=0
        Do While Not ti.AtEndOfStream
            anyTmp=Split(ti.ReadLine,",")
            LineNo=LineNo+1
            If anyTmp(1)=strFindCntnr Then
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

    ' 輸入コンテナ照会
'    WriteLog fs, "輸入コンテナ照会", "画面更新"

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
          <td rowspan=2><img src="gif/csvt.gif" width="506" height="73"></td>
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
          <td nowrap>
            <dl> 
            <dt><font color="#000066" size="+1">(Screen for file transfer for container information inquiry)</font><br>
            <dd>
<%
    ' エラーメッセージの表示
    DispErrorMessage strError
%>
            </dl>
          </td>
        </tr>
      </table>
      <form action="impentry.asp">
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
    DispMenuBarBack "impentry.asp"
%>
</body>
</html>

<%
    Else
        If strRequest="impdetail.asp" Then
            ' 詳細画面へリダイレクト
            Response.Redirect "impdetail.asp?line=" & LineNo  '輸入コンテナ詳細
        Else
            ' 一覧画面へリダイレクト
            Response.Redirect strRequest                      '輸入コンテナ一覧
        End If
    End If
%>
