<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="expcom.inc"-->

<%
''    ' セッションのチェック
''    CheckLogin "expentry.asp"

    '入力画面を記憶
    Session.Contents("findcsv")="yes"    ' CSVファイル入力であることを記憶

    ' 指定引数の取得
    Dim strKind
    strKind = "booking"

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' 参照モードをセッション変数に設定
    Session.Contents("findkind")="Booking"

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

		Dim sOutText()
		Dim strOut,bWrite

        For iCount=0 To iRecCount - 1
            sWhere = "Booking.BookNo='" & strCntnrNo(iCount) & "'"

'	        sql = "SELECT Booking.BookNo," &_
'						  "Booking.VslCode," &_
'						  "Booking.ShipLine," &_
'						  "mShipLine.FullName ShipLineName," &_
'						  "mVessel.FullName ShipName," &_
'						  "VslSchedule.LdVoyage," &_
'						  "Booking.DPort," &_
'						  "mPort.FullName PortName," &_
'						  "Pickup.PickPlace," &_
'						  "Pickup.ContSize," &_
'						  "Pickup.ContType," &_
'						  "Pickup.ContHeight," &_
'						  "Pickup.Qty " &_
'				  "FROM Booking,mShipLine,mVessel,VslSchedule,Pickup,mPort " &_
'				  "WHERE (" & sWhere & ") AND " &_
'						  "Booking.VslCode=Pickup.VslCode AND " &_
'						  "Booking.VoyCtrl=Pickup.VoyCtrl AND " &_
'						  "Booking.BookNo=Pickup.BookNo AND " &_
'						  "mShipLine.ShipLine=*Booking.ShipLine AND " &_
'						  "mVessel.VslCode=*Booking.VslCode AND " &_
'						  "mPort.PortCode=*Booking.DPort AND " &_
'						  "VslSchedule.VslCode=*Booking.VslCode AND " &_
'						  "VslSchedule.VoyCtrl=*Booking.VoyCtrl" 		'D20040223

	        sql = "SELECT Booking.BookNo," &_
						  "Booking.VslCode," &_
						  "Booking.ShipLine," &_
						  "mShipLine.FullName ShipLineName," &_
						  "mVessel.FullName ShipName," &_
						  "VslSchedule.LdVoyage," &_
						  "Booking.DPort," &_
						  "mPort.FullName PortName," &_
						  "Pickup.PickPlace," &_
						  "Pickup.ContSize," &_
						  "Pickup.ContType," &_
						  "Pickup.ContHeight," &_
						  "Pickup.Qty," &_
						  "Pickup.Material " &_
				  "FROM Booking,mShipLine,mVessel,VslSchedule,Pickup,mPort " &_
				  "WHERE (" & sWhere & ") AND " &_
						  "Booking.VslCode=Pickup.VslCode AND " &_
						  "Booking.VoyCtrl=Pickup.VoyCtrl AND " &_
						  "Booking.BookNo=Pickup.BookNo AND " &_
						  "mShipLine.ShipLine=*Booking.ShipLine AND " &_
						  "mVessel.VslCode=*Booking.VslCode AND " &_
						  "mPort.PortCode=*Booking.DPort AND " &_
						  "VslSchedule.VslCode=*Booking.VslCode AND " &_
						  "VslSchedule.VoyCtrl=*Booking.VoyCtrl" 		'I20040223

			rsd.Open sql, conn, 0, 1, 1

		    bWrite = 0        '出力レコード件数

			Do While Not rsd.EOF
				strOut = Trim(rsd("VslCode")) & ","						' 0:VslCode
				strOut = strOut & Trim(rsd("BookNo")) & ","				' 1:Booking No.

				If IsNull(rsd("ShipLineName")) Then
					strOut = strOut & Trim(rsd("ShipLine")) & ","		' 2:船社
				Else
					strOut = strOut & Trim(rsd("ShipLineName")) & ","	' 2:船社
				End If

				If IsNull(rsd("ShipName")) Then
					strOut = strOut & Trim(rsd("VslCode")) & ","		' 3:船名
				Else
					strOut = strOut & Trim(rsd("ShipName")) & ","		' 3:船名
				End If

				strOut = strOut & Trim(rsd("LdVoyage")) & ","			' 4:Voyage No.

				If IsNull(rsd("PortName")) Then
					strOut = strOut & Trim(rsd("DPort")) & ","			' 5:仕向港
				Else
					strOut = strOut & Trim(rsd("PortName")) & ","		' 5:仕向港
				End If

				strOut = strOut & Trim(rsd("PickPlace")) & ","			' 6:ピック場所
				strOut = strOut & Trim(rsd("ContSize")) & ","			' 7:サイズ
				strOut = strOut & Trim(rsd("ContType")) & ","			' 8:タイプ
				strOut = strOut & Trim(rsd("ContHeight")) & ","			' 9:高さ
				strOut = strOut & Trim(rsd("Qty")) & ","				'10:予約本数

				strOut = strOut & "," & Trim(rsd("Material"))			'12:材質	'I20040223

				ReDim Preserve sOutText(bWrite)
				sOutText(bWrite) = strOut
				bWrite = bWrite + 1

				rsd.MoveNext
			Loop

		    rsd.Close

		    For i=0 To bWrite-1
		        strTmp=Split(sOutText(i),",")

'				sql = "SELECT ExportCont.ContNo FROM ExportCont,Container " &_
'					  "WHERE ExportCont.VslCode='" & strTmp(0) & "'" &_
'					   " AND ExportCont.BookNo='" & strTmp(1) & "'" &_
'					   " AND ExportCont.PickPlace='" & strTmp(6) & "'" &_
'					   " AND Container.VslCode='" & strTmp(0) & "'" &_
'					   " AND ExportCont.VoyCtrl=Container.VoyCtrl" &_
'					   " AND ExportCont.ContNo=Container.ContNo" &_
'					   " AND Container.ContSize='" & strTmp(7) & "'" &_
'					   " AND Container.ContType='" & strTmp(8) & "'" &_
'					   " AND Container.ContHeight='" & strTmp(9) & "'" 		'D20040223

				sql = "SELECT ExportCont.ContNo FROM ExportCont,Container " &_
					  "WHERE ExportCont.VslCode='" & strTmp(0) & "'" &_
					   " AND ExportCont.BookNo='" & strTmp(1) & "'" &_
					   " AND ExportCont.PickPlace='" & strTmp(6) & "'" &_
					   " AND Container.VslCode='" & strTmp(0) & "'" &_
					   " AND ExportCont.VoyCtrl=Container.VoyCtrl" &_
					   " AND ExportCont.ContNo=Container.ContNo" &_
					   " AND Container.ContSize='" & strTmp(7) & "'" &_
					   " AND Container.ContType='" & strTmp(8) & "'" &_
					   " AND Container.ContHeight='" & strTmp(9) & "'" &_
				   " AND Container.Material='" & strTmp(12) & "'"		'I20040223

				rsd.Open sql, conn, 0, 1, 1

				Dim iContNum
				iContNum = 0
				sOutText(i) = sOutText(i) & ","

	            Do While Not rsd.EOF
					If iContNum=0 Then
						sOutText(i) = sOutText(i) & Trim(rsd("ContNo"))
					Else
						sOutText(i) = sOutText(i) & "/" & Trim(rsd("ContNo"))
					End If
					iContNum = iContNum + 1
					rsd.MoveNext
				Loop

		        rsd.Close

	     		strTmpIn=Split(sOutText(i),",")
				strTmpIn(11) = iContNum					'11:搬出済本数
				sOutTextTmp = strTmpIn(0)
				For k=1 To UBound(strTmpIn)
					sOutTextTmp = sOutTextTmp & "," & strTmpIn(k)
				Next

				ti.WriteLine sOutTextTmp
			Next

			bWriteFile = bWriteFile + bWrite
        Next

        ti.Close
        conn.Close

        If bWriteFile = 0 Then
            ' 該当レコードないとき
            bError = true
            strError = "指定条件に該当するBooking No.はありませんでした。"
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

    ' 輸出コンテナ照会
    WriteLog fs, "1012","ブッキング情報照会-CSVファイル転送","20", strOption

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
          <td rowspan=2><img src="gif/bookingt.gif" width="506" height="73"></td>
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
end of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>CSVファイル転送</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr> 
          <td nowrap>
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
    DispMenuBarBack "bookcsv.asp"
%>
</body>
</html>

<%
    Else
        ' 一覧画面へリダイレクト
        Response.Redirect "booklist.asp"
    End If
%>
