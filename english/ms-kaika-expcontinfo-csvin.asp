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
	Const sTranID05 = "EX05"
	Const sTranID16 = "EX16"
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
			If Ubound(anytmp) <> 10 Then
                ' ファイル形式エラー
                strError="項目数が異常です。"
			Else
				'各項目桁数＆整合性チェック
				'船名 anyTmp(0)
				strError = strError & CheckParam(anyTmp(0),"船名",7,0,true,false) 
 				'Voyage No. anyTmp(1)
				strError = strError & CheckParam(anyTmp(1),"Voyage No.",12,0,true,false) 
 				'荷主コード anyTmp(2)
				strError = strError & CheckParam(anyTmp(2),"荷主コード",5,0,true,false) 
 				'荷主管理番号 anyTmp(3)
				strError = strError & CheckParam(anyTmp(3),"荷主管理番号",10,0,true,false) 
 				'Booking No. anyTmp(4)
				strError = strError & CheckParam(anyTmp(4),"Booking No.",20,0,true,false) 
 				'コンテナNo. anyTmp(5)
				strError = strError & CheckParam(anyTmp(5),"コンテナNo.",12,0,true,false) 
 				'シールNo. anyTmp(6)
				strError = strError & CheckParam(anyTmp(6),"シールNo.",15,0,false,false) 
 				'貨物重量 anyTmp(7)
				strError = strError & CheckParam(anyTmp(7),"貨物重量",4,0,false,true) 
 				'総重量 anyTmp(8)
				strError = strError & CheckParam(anyTmp(8),"総重量",4,0,false,true) 
 				'リーファー anyTmp(9)
				strError = strError & CheckParam(anyTmp(9),"リーファー",1,0,false,true) 
 				'危険物 anyTmp(10)
				strError = strError & CheckParam(anyTmp(10),"危険物",1,0,false,true) 

                ' リーファー／危険物のチェック
                sTemp=Trim(anyTmp(9))
                If sTemp<>"" And sTemp<>"1" And sTemp<>"0" Then
                    strError=strError & "リーファーが異常です。(" & anyTmp(9) & ") "
                End If
                sTemp=Trim(anyTmp(10))
                If sTemp<>"" And sTemp<>"1" And sTemp<>"0" Then
                    strError=strError & "危険物が異常です。(" & anyTmp(10) & ") "
                End If

				'エラー時にSQLをなるべく発行しないようにIf文で括る
				If strError="" Then
					' コンテナNo.が存在するか
					Dim sVanTime,sShipLine
					sql = "SELECT ExportCont.VanTime,VslSchedule.ShipLine " & _
						  "FROM ExportCont,VslSchedule " & _
						  "WHERE " & _
							"ExportCont.VslCode='" & anyTmp(0) & "' AND " & _
							"ExportCont.ContNo='" & anyTmp(5) & "' AND " & _
							"ExportCont.BookNo='" & anyTmp(4) & "' AND " & _
							"VslSchedule.VslCode='" & anyTmp(0) & "'"
					rsd.Open sql, conn, 0, 1, 1
					If Not rsd.EOF Then
'					    sVanTime  = Trim(rsd("VanTime"))
					    sShipLine = Trim(rsd("ShipLine"))
					Else
'						strError = "指定されたコンテナNo.が存在しません。(" & anyTmp(5) & ") "
					End If
					rsd.Close

					If anyTmp(5) = "" Then
						strError = "コンテナNo.が指定されていません。(" & anyTmp(5) & ") "
					End If

                End If

				If strError="" Then
					'CSVファイルの行をTmpに格納
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
			'ExportCargoInfoにコンテナNo.が入っているか
			Dim sCont,sContSize,sContType,sContHeight,sRemark,sTrucker,sWHArTime,sCYRecDate,sPickPlace
			sql = "SELECT ContNo,ContSize,ContType,ContHeight,Remark,Trucker,WHArTime,CYRecDate,PickPlace " & _
				  "FROM ExportCargoInfo " & _
				  "WHERE Shipper='" &  UCase(Trim(anyTmp(2))) & _
					"' And ShipCtrl='" &  UCase(Trim(anyTmp(3))) & "'"
			rsd.Open sql, conn, 0, 1, 1
			If Not rsd.EOF Then
			    sCont  		= Trim(rsd("ContNo"))
			    sContSize 	= Trim(rsd("ContSize"))
			    sContType 	= Trim(rsd("ContType"))
			    sContHeight = Trim(rsd("ContHeight"))
			    sRemark 	= Trim(rsd("Remark"))
			    sTrucker 	= Trim(rsd("Trucker"))
			    sWHArTime 	= Trim(rsd("WHArTime"))
			    sCYRecDate 	= Trim(rsd("CYRecDate"))
			    sPickPlace 	= Trim(rsd("PickPlace"))
			Else
				strError = "荷主コード、荷主管理番号が異常です。(" & anyTmp(2) & "," & anyTmp(3) & ") "
			End If
			rsd.Close

			'コンテナNo.が空またはCSVと異なる場合EX16を作成
			Dim sEX05, iSeqNo_EX05, sEX16, iSeqNo_EX16, sFileName, strFileName_01, sTran, sTusin
			iOutCount = 0

			If sCont="" Or sCont<>UCase(Trim(anyTmp(5))) Then
	            ' 出力ファイル設定
				iSeqNo_EX16 = GetDailyTransNo

				sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX16
				strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
			    Set tout=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

	            For iCount=0 To iWriteCnt - 1
	                'シーケンス番号
	                anyTmp1 = Split(Tmp(iCount), ",")
					If iCount <> 0  Then
						iSeqNo_EX16 = GetDailyTransNo
					End If

'トランザクション作成時CSVファイル内項目にTrimとUCaseをかける  2002/02/04
					For j=0 To 10
						anyTmp1(j) = UCase(Trim(anyTmp1(j)))
					Next
'ここまで
					'通信日時取得
					sTusin  = SetTusinDate

					sEX16 = iSeqNo_EX16 & "," & sTranID16 & "," & sSyori & ","  & sTusin & ",Web - " & _
							sSosin & "," & sPlace & "," & anyTmp1(0) & "," &  anyTmp1(1) & "," & _
							anyTmp1(4) & "," & anyTmp1(2) & "," & anyTmp1(3) & "," & sSosin & "," & _
							anyTmp1(5) & "," & sContSize & "," & sContType & "," & sContHeight & "," & _
							sRemark & "," & sTrucker & "," & _
							sWHArTime & "," & sCYRecDate & "," & sPickPlace
					tout.WriteLine sEX16
	                iOutCount=iOutCount+1
				Next 

			    tout.Close
			End If

            ' 出力ファイル設定
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

'トランザクション作成時CSVファイル内項目にTrimとUCaseをかける  2002/02/04
				For j=0 To 10
					anyTmp1(j) = UCase(Trim(anyTmp1(j)))
				Next
'ここまで
				If anyTmp1(9)=1 And anyTmp1(10)=1 Then
					anyTmp1(9) = "RH"
				ElseIf anyTmp1(9)=1 Then
					anyTmp1(9) = "R"
				ElseIf anyTmp1(10)=1 Then
					anyTmp1(9) = "H"
				Else
					anyTmp1(9) = ""
				End If

				'通信日時取得
				sTusin  = SetTusinDate

				sEX05 = iSeqNo_EX05 & "," & sTranID05 & "," & sSyori & ","  & sTusin & ",Web - " & _
						sSosin & "," & sPlace & "," & anyTmp1(0) & "," &  anyTmp1(1) & "," & _
						anyTmp1(5) & "," & anyTmp1(4) & "," & sVanTime & "," & sShipLine & "," & _
						anyTmp1(8)*10 & "," & anyTmp1(6) & "," & anyTmp1(7)*10 & "," & _
						sSosin & ",," & anyTmp1(9)
				tout.WriteLine sEX05
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

    WriteLog fs, "4107","海貨入力輸出コンテナ情報-CSVファイル転送","20", strOption

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
          <td rowspan=2><img src="gif/kaika5t.gif" width="506" height="73"></td>
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
            <font color="#000066" size="+1">【輸出コンテナ情報用ファイル転送画面】</font><BR><br>
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
    DispMenuBarBack "ms-kaika-expcontinfo-csv.asp"
%>
</body>
</html>

