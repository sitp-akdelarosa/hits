<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->
<!--#include file="csvcheck.inc"-->

<%
''    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"

	sSosin = Trim(Session.Contents("userid"))

	' トランザクションファイルの拡張子 
	Const SEND_EXTENT = "snd"
	' トランザクションＩＤ
	Const sTranID = "IM18"
	' 処理区分
	Const sSyori = "R"

	' 送信場所
	Const sPlace = ""
    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile") = strFileName

    ' 転送ファイルの取得
    tb=Request.TotalBytes      :' ブラウザからのトータルサイズ
    br=Request.BinaryRead(tb)  :' ブラウザからの生データ

    ' BASP21 コンポーネントの作成
    Set bsp=Server.CreateObject("basp21")

    filesize=bsp.FormFileSize(br,"csvfile")
    filename=bsp.FormFileName(br,"csvfile")

    fpath=strFileName
    fpath=fs.BuildPath(Server.MapPath("./temp"),fpath)

    lng=bsp.FormSaveAs(br,"csvfile",fpath)

    ' ファイル転送に失敗したとき

    Dim sText	'転送ファイル

    If lng<=0 Then
        bError=true
        strError = "'" & filename & "'ファイルの転送に失敗しました。"
    Else
        ' 転送ファイルのOpen
        Set ti=fs.OpenTextFile(fpath,1,True)
		Dim anyTmp, iRecCount, iWriteCnt, iErrLine
		iRecCount = 0
		iWriteCnt = 0
		iErrLine = 0

        ConnectSvr conn, rsd
        ' 転送ファイルのレコードがある間繰り返す
        Do While Not ti.AtEndOfStream
            strError=""
			sText = ti.ReadLine
			anyTmp = Split(sText, ",")
			If Ubound(anytmp) <> 9 Then
                ' ファイル形式エラー
                strError="項目数が異常です。"
			Else
				'各項目桁数＆整合性チェック
				'船名 anyTmp(0)
				strError = strError & CheckParam(anyTmp(0),"船名(コールサイン)",7,0,true,false) 
 				'Voyage No. anyTmp(1)
				strError = strError & CheckParam(anyTmp(1),"Voyage No.",12,0,true,false) 
 				'荷主コード anyTmp(2)
				strError = strError & CheckParam(anyTmp(2),"荷主コード",5,0,true,false) 
 				'BL No. anyTmp(3)
				strError = strError & CheckParam(anyTmp(3),"BL No.",20,0,true,false) 
 				'コンテナNo. anyTmp(4)
				strError = strError & CheckParam(anyTmp(4),"コンテナNo.",12,0,true,false) 
 				'指定陸運業者コード anyTmp(5)
				strError = strError & CheckParam(anyTmp(5),"指定陸運業者コード",3,0,false,false) 
 				'実入り倉庫到着予定日時 anyTmp(6)
				'日付のスラッシュを取って桁数をあわせる
				sTemp=ChangeDate(Trim(anyTmp(6)),12)
           	    If InStr(sTemp,"(")<>0 Then
                    ' 入力データ エラー
                    strError=strError & "実入り倉庫到着予定日時の" & sTemp
                End If
 				'サイズ anyTmp(7)
				strError = strError & CheckParam(anyTmp(7),"サイズ",2,0,false,true) 
 				'タイプ anyTmp(8)
				strError = strError & CheckParam(anyTmp(8),"タイプ",2,0,false,false) 
 				'倉庫略称 anyTmp(9)
				strError = strError & CheckParam(anyTmp(9),"倉庫略称",5,0,false,false) 

				'エラー時にSQLをなるべく発行しないようにIf文で括る
				If strError="" Then
					Dim iRecCnt
					' 船名と次航とコンテナNo.の重複チェック
					If Not bKind=0 Then
						sql = "SELECT count(*) FROM ImportCargoInfo " & _
								"WHERE VslCode='" & UCase(Trim(anyTmp(0))) & _
								"' AND DsVoyage='" & UCase(Trim(anyTmp(1))) & _
								"' AND ContNo='" & UCase(Trim(anyTmp(4))) & "'"
						rsd.Open sql, conn, 0, 1, 1
						If Not rsd.EOF Then
							iRecCnt = rsd(0)
							If Not iRecCnt=0 Then
								strError = "船名, VoyageNo, コンテナNo.が重複しています。(" & anyTmp(0) & "," & anyTmp(1) & "," & anyTmp(4) & ") "
							End If
						End If
						rsd.Close
					End If

					' 船名が存在するか
					sql = "SELECT count(*) FROM mVessel WHERE VslCode='" &  UCase(Trim(anyTmp(0))) & "'"
					rsd.Open sql, conn, 0, 1, 1
					If Not rsd.EOF Then
						iRecCnt = rsd(0)
						If iRecCnt=0 Then
							strError = "指定された船名が存在しません。(" & anyTmp(0) & ") "
						End If
					End If
					rsd.Close
				End If

                If strError="" Then
                    ReDim Preserve Tmp(iWriteCnt)
                    Tmp(iWriteCnt) = sText
                    iWriteCnt = iWriteCnt + 1
                End If
            End If
            iRecCount = iRecCount + 1
            If strError<>"" Then
                ReDim Preserve sErrLine(iErrLine)
                sErrLine(iErrLine) = iRecCount & "件目:" & strError
                iErrLine = iErrLine + 1
            End If
        Loop
        ti.Close

        If iErrLine > 0 Then
            bError = true
            strError = "'" & filename & "'ファイルの形式が違います。" & "<br>"
            For i = 0 to iErrLine - 1
                strError = strError & sErrLine(i) & "<br>"
            Next
        Else
            iOutCount=0
            ' 出力ファイル設定
			Dim sIM18, iSeqNo_IM18, sFileName, strFileName_01, sTran, sTusin
			iSeqNo_IM18 = GetDailyTransNo

			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM18
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		    Set tout=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

            For iCount=0 To iWriteCnt - 1
                'シーケンス番号
                anyTmp1 = Split(Tmp(iCount), ",")
				If iCount <> 0  Then
					iSeqNo_IM18 = GetDailyTransNo
				End If

'トランザクション作成時CSVファイル内項目にTrimとUCaseをかける  2002/02/04
				For j=0 To 9
					anyTmp1(j) = UCase(Trim(anyTmp1(j)))
				Next
'ここまで

				'日付のスラッシュを取って桁数をあわせる
				anyTmp1(6)=ChangeDate(Trim(anyTmp1(6)),12)

				'通信日時取得
				sTusin  = SetTusinDate

				sIM18 = iSeqNo_IM18 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
						sSosin & "," & sPlace & "," & anyTmp1(0) & "," &  anyTmp1(1) & "," & _
						anyTmp1(3) & "," & anyTmp1(2) & "," & sSosin & "," & _
						anyTmp1(4) & "," & anyTmp1(7) & "," & anyTmp1(8) & "," & _
						anyTmp1(5) & "," & anyTmp1(9) & "," & anyTmp1(6)
				tout.WriteLine sIM18
                iOutCount=iOutCount+1
			Next 

		    tout.Close

		    ' エラーメッセージの表示
			strError = "正常に更新されました。"
		End IF
    End If

	' Logファイル書き出し
    If bError Then
        strOption = filename & "," & "入力内容の正誤:1(誤り)"
    Else
        strOption = filename & "," & "入力内容の正誤:0(正しい) " & iOutCount & "件出力"
    End If

    WriteLog fs, "4111","海貨入力輸入コンテナ情報-CSVファイル転送","20", strOption

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
          <td rowspan=2><img src="gif/kaika6t.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>CSVファイル転送</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr> 
          <td nowrap align=center>
            <font color="#000066" size="+1">【輸入コンテナ情報用ファイル転送画面】</font><br><BR>
<%
    ' エラーメッセージの表示
    If strError="正常に更新されました。" Then
        DispInformationMessage strError
    Else
        DispErrorMessage strError
    End If
%>
            </dl>
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
    DispMenuBarBack "ms-kaika-impcontinfo-csv.asp"
%>
</body>
</html>

