<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->

<%
''    ' セッションのチェック
    CheckLogin "nyuryoku-ki.asp"

	sSosin = Trim(Session.Contents("userid"))

	' トランザクションファイルの拡張子 
	Const SEND_EXTENT = "snd"
	' トランザクションＩＤ
	Const sTranID = "EX05"
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
			If Ubound(anytmp) <> 4 Then
                ' ファイル形式エラー
                strError="項目数が異常です。"
			Else
				'ファイル形式的には正常
                ' 入力コンテナNo.のチェック
                sql = "SELECT ExportCont.VslCode, ExportCont.VoyCtrl, ExportCont.BookNo, ExportCont.WHArTime, VslSchedule.LdVoyage, VslSchedule.ShipLine "
                sql = sql & " FROM ExportCont, VslSchedule"
                sql = sql & " WHERE ExportCont.ContNo='" & Trim(anyTmp(0)) & "' And VslSchedule.VslCode = ExportCont.VslCode"
                sql = sql & " AND VslSchedule.VoyCtrl = ExportCont.VoyCtrl"

                'SQLを発行して輸出コンテナを検索
                rsd.Open sql, conn, 0, 1, 1
                If Not rsd.EOF Then
                    sVslCode = Trim(rsd("VslCode"))		'船名
                    sVoyCtrl = Trim(rsd("LdVoyage"))	'次航
                    sBookNo = Trim(rsd("BookNo"))		'ブッキング
                    stShipLine = Trim(rsd("ShipLine"))	'船社
'                   stWHArTime = GetYMDHM(rsd("WHArTime")) 		'バン詰め日時
                    sText=sVslCode & "," & sVoyCtrl & "," & Trim(anyTmp(0)) & "," & sBookNo & "," & stShipLine & "," & stWHArTime
                Else
                    ' コンテナ エラー
                    strError=strError & "該当するコンテナが存在しません。(" & anyTmp(0) & ") "
                End If
                rsd.Close
                ' シールNo.のチェック
                If Len(Trim(anyTmp(1)))>15 Or Len(Trim(anyTmp(1)))<=0 Then
                    ' シールNo.の長さ エラー
                    strError=strError & "シールNo.の長さが異常です。(" & anyTmp(1) & ") "
                Else
                    sText=sText & "," & Trim(anyTmp(1))
                End If
                ' 貨物重量のチェック
                If Trim(anyTmp(2))<>"" Then
                    If IsNumeric(Trim(anyTmp(2))) Then
                        fTemp=CDbl(Trim(anyTmp(2)))
                        If fTemp>99.9 Or fTemp<0 Then
                            ' 貨物重量 エラー
                            strError=strError & "貨物重量は99.9Tonまでです。(" & anyTmp(2) & ") "
                        Else
                            sText=sText & "," & CInt(fTemp*10)
                        End If
                    Else
                        ' 貨物重量 エラー
                        strError=strError & "貨物重量が異常です。(" & anyTmp(2) & ") "
                    End If
                Else
                    sText=sText & ","
                End If
                ' 総重量のチェック
                If Trim(anyTmp(3))<>"" Then
                    If IsNumeric(Trim(anyTmp(3))) Then
                        fTemp=CDbl(Trim(anyTmp(3)))
                        If fTemp>99.9 Or fTemp<0 Then
                            ' 総重量 エラー
                            strError=strError & "総重量は99.9Tonまでです。(" & anyTmp(3) & ") "
                        Else
                            sText=sText & "," & CInt(fTemp*10)
                        End If
                    Else
                        ' 総重量 エラー
                        strError=strError & "総重量が異常です。(" & anyTmp(3) & ") "
                    End If
                Else
                    sText=sText & ","
                End If
                ' リーファー／危険物のチェック
                sTemp=Trim(anyTmp(4))
                If sTemp<>"" And sTemp<>"R" And sTemp<>"H" And sTemp<>"RH" And sTemp<>"HR" Then
                    ' リーファー／危険物 エラー
                    strError=strError & "リーファー／危険物が異常です。(" & anyTmp(4) & ") "
                Else
                    sText=sText & "," & sTemp
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
			Dim sEX05, iSeqNo_EX05, sFileName, strFileName_01, sTran, sTusin
			iSeqNo_EX05 = GetDailyTransNo

			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX05
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		    Set tout=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

            For iCount=0 To iWriteCnt - 1
                'シーケンス番号
                anyTmp1 = Split(Tmp(iCount), ",")
				If iCount <> 0  Then
					iSeqNo_EX05 = GetDailyTransNo
				End If
				'通信日時取得
				sTusin  = SetTusinDate

				sEX05 = iSeqNo_EX05 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
						sSosin & "," & sPlace & "," & anyTmp1(0) & "," &  anyTmp1(1) & "," & _
						anyTmp1(2) & "," & anyTmp1(3) & "," & anyTmp1(4) & "," & anyTmp1(5) & "," & _
						anyTmp1(8) & "," & anyTmp1(6) & "," & anyTmp1(7) & "," & sSosin & ",," & anyTmp1(9)
				tout.WriteLine sEX05
                iOutCount=iOutCount+1
			Next 

		    tout.Close

		    ' エラーメッセージの表示
			strError = "正常に更新されました。"
		End IF
    End If

    If bError Then
        strOption = filename & "," & "入力内容の正誤:1(誤り)"
    Else
        strOption = filename & "," & "入力内容の正誤:0(正しい) " & iOutCount & "件出力"
    End If

    ' 海貸用ファイル転送画面照会
    WriteLog fs, "4005","海貨入力シールNo.・重量入力-CSVファイル転送","20", strOption

'''    If bError Then
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
            <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
            <td nowrap><b>（輸出）シールNo.・重量</b></td>
            <td><img src="gif/hr.gif"></td>
          </tr>
		</table>
      <table>
        <tr> 
          <td nowrap align=center>
            <font color="#000066" size="+1">【シールNo.・重量用ファイル転送画面】</font><br><BR>
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
    DispMenuBarBack "nyuryoku-kcsv.asp"
%>
</body>
</html>

