<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	【コンテナ情報入力】	更新時表示チェック、Tempファイル作成
	'
%>

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"

	' 検索一覧表示最大値
	Dim sUser,sUserNo
    sUser    = UCase(Trim(Request.form("contuser")))
    sUserNo  = UCase(Trim(Request.form("contuserno")))
    sBooking = UCase(Trim(Request.form("contbooking")))

	' 検索一覧表示最大値
	Dim iMaxCount
	iMaxCount = 100
    ' エラーフラグのクリア
    bError = false
	' 海貨コード取得
	sSosin = Trim(Session.Contents("userid"))

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' レコード数の取得
	Dim iRecCount,strErrMsg
    ConnectSvr conn, rsd
	sql = "SELECT count(*) FROM ExportCargoInfo WHERE Forwarder='" & sSosin & "'"
	If sUser<>"" Then
		sql = sql & " AND Shipper='" & sUser & "'"
	End If
	If sUserNo<>"" Then
		sql = sql & " AND ShipCtrl='" & sUserNo & "'"
	End If
	If sBooking<>"" Then
		sql = sql & " AND BookNo='" & sBooking & "'"
	End If

	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
	    iRecCount = rsd(0)
	Else
	    bError = true
		strErrMsg = "DB接続エラー"
	End If
	rsd.Close

	If iRecCount>iMaxCount Then
	    bError = true
		strErrMsg = "検索対象件数が最大値を超えています。<BR>絞り込みをして下さい。"
	Else If iRecCount=0 Then
	    bError = true
		strErrMsg = "対象データが存在しません。"
	Else
		Dim strOut,bWrite
		' Tempファイル書き出し
	    bWrite = 0        '出力レコード件数

	    ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
	    strFileName="./temp/" & strFileName
	    ' テンポラリファイルのOpen
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

		sql = "SELECT ExportCargoInfo.Shipper,ExportCargoInfo.ShipCtrl,ExportCargoInfo.BookNo," & _
				"ExportCargoInfo.ContNo,ExportCargoInfo.VslCode,ExportCargoInfo.LdVoyage," & _
				"ExportCont.RFFlag, ExportCont.DGFlag," & _
				"Container.ContWeight,Container.SealNo,Container.CargoWeight " & _
			  "FROM ExportCargoInfo,Container,ExportCont " & _
			  "WHERE ExportCargoInfo.Forwarder='" & sSosin & "' And "

		If sUser<>"" Then
			sql = sql & "ExportCargoInfo.Shipper='" & sUser & "' And "
		End If
		If sUserNo<>"" Then
			sql = sql & "ExportCargoInfo.ShipCtrl='" & sUserNo & "' And "
		End If
		If sBooking<>"" Then
			sql = sql & "ExportCargoInfo.BookNo='" & sBooking & "' And "
		End If

		sql = sql & "ExportCont.VslCode=*ExportCargoInfo.VslCode And " & _
					"ExportCont.ContNo=*ExportCargoInfo.ContNo And " & _
					"ExportCont.BookNo=*ExportCargoInfo.BookNo And " & _
					"Container.VslCode=*ExportCargoInfo.VslCode And " & _
					"Container.ContNo=*ExportCargoInfo.ContNo"

	    rsd.Open sql, conn, 0, 1, 1

	    Do While Not rsd.EOF
	        strOut = Trim(rsd("VslCode")) & ","
     		strOut = strOut & Trim(rsd("LdVoyage")) & ","
     		strOut = strOut & Trim(rsd("Shipper")) & ","
     		strOut = strOut & Trim(rsd("ShipCtrl")) & ","
     		strOut = strOut & Trim(rsd("BookNo")) & ","
     		strOut = strOut & Trim(rsd("ContNo")) & ","
     		strOut = strOut & Trim(rsd("SealNo")) & ","
     		strOut = strOut & Trim(rsd("CargoWeight"))/10 & ","
     		strOut = strOut & Trim(rsd("ContWeight"))/10 & ","
     		strOut = strOut & Trim(rsd("RFFlag")) & ","
     		strOut = strOut & Trim(rsd("DGFlag"))

	        ti.WriteLine strOut
	        bWrite = bWrite + 1

	        rsd.MoveNext
	    Loop

  		rsd.Close

	End If
	End If

    If bError Then
        strOption = sUser & "/" & sUserNo & "/" & sBooking & "," & "入力内容の正誤:1(誤り)"
    Else
        strOption = sUser & "/" & sUserNo & "/" & sBooking & "," & "入力内容の正誤:0(正しい)"
    End If

    WriteLog fs, "4105","海貨入力輸出コンテナ情報","10", strOption

	If Not bError Then
        Response.Redirect "ms-kaika-expcontinfo-list.asp"
	Else
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
<!-------------ここからログイン入力画面--------------------------->
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>更新対象一覧</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>
<%
    ' エラーメッセージの表示
    DispErrorMessage strErrMsg 
%>
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
<!-------------ログイン画面終わり--------------------------->
<%
    DispMenuBarBack "ms-kaika-expcontinfo.asp"
%>
</body>
</html>

<%
	End If
%>
