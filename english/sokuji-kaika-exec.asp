<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	即時搬出システム【海貨用】	変更,削除用画面
%>

<%
	' セッションのチェック
	CheckLogin "sokuji.asp"

	' 海貨コード
	sForwarder = Trim(Session.Contents("userid"))

	Dim bKind,sSend,sStop,sDel,iLineNo
	' 新規登録後(2) or 新規(1) or 更新(0)
    bKind = Trim(Session.Contents("kind"))
	' 種別
    sSend 	= Trim(Request.form("send"))
    sStop 	= Trim(Request.form("stop"))
    sDel 		= Trim(Request.form("del"))
    iLineNo	= Trim(Request.form("lineno"))

	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	If bKind=0 And sStop<>"" Then
		Response.Redirect "sokuji-kaika-list.asp"
		Response.End
	ElseIf bKind=1 And sStop<>"" Then
		Response.Redirect "sokuji-kaika-list.asp"
		Response.End
	ElseIf bKind=2 And sStop<>"" Then
		Response.Redirect "sokuji-kaika-list.asp"
		Response.End
	Else
		' トランザクションファイルの拡張子 
		Const SEND_EXTENT = "snd"
		' トランザクションＩＤ
		Const sTranID = "IM19"
		' 送信場所
		Const sPlace = ""
		' エラーフラグのクリア
		bError = false

		' 指定引数の取得
		Dim sShipper,sShipLine,sVslCode,sBL,sCont,sOpe,sOpeTel,sBLold,sContold',sReject,sRecschTime
		Dim sTmpVslCode,sTmpVoyCtrl
		sShipper 	= UCase(Trim(Request.form("shipper")))
		sShipLine = UCase(Trim(Request.form("shipline")))
		sVslCode = UCase(Trim(Request.form("vslcode")))
		sBL = UCase(Trim(Request.form("bl")))
		sCont = UCase(Trim(Request.form("cont")))
'		sOpe 	= UCase(Trim(Request.form("ope")))
'		sOpeTel 	= UCase(Trim(Request.form("opetel")))
		If sBL<>"" Then
			ConnectSvr conn, rsd
			sql = "SELECT mOperator.NameAbrev,mOperator.TelNo FROM BL,mOperator " & _
				"WHERE BLNo='" & sBL & "' AND BL.OpeCode=mOperator.OpeCode"
			rsd.Open sql, conn, 0, 1, 1
			Do While Not rsd.EOF
				sOpe = Trim(rsd(0))
				sOpeTel = Trim(rsd(1))
				Exit Do
				rsd.MoveNext
			Loop
			rsd.Close
		Else
			ConnectSvr conn, rsd
			sql = "SELECT VslCode,VoyCtrl FROM ImportCont " & _
				"WHERE ContNo='" & sCont & "' ORDER BY UpdtTime DESC"
			rsd.Open sql, conn, 0, 1, 1
			Do While Not rsd.EOF
				sTmpVslCode = Trim(rsd(0))
				sTmpVoyCtrl = Trim(rsd(1))
				Exit Do
				rsd.MoveNext
			Loop
			rsd.Close
			sql = "SELECT mOperator.NameAbrev,mOperator.TelNo FROM BL,mOperator " & _
				"WHERE BL.VslCode='" & sTmpVslCode & "' AND BL.VoyCtrl='" & sTmpVoyCtrl & _
				"' AND BL.OpeCode=mOperator.OpeCode"
			rsd.Open sql, conn, 0, 1, 1
			Do While Not rsd.EOF
				sOpe = Trim(rsd(0))
				sOpeTel = Trim(rsd(1))
				Exit Do
				rsd.MoveNext
			Loop
			rsd.Close
		End If
		sBLold = UCase(Trim(Request.form("blold")))
		sContold = UCase(Trim(Request.form("contold")))
'		sReject 	= UCase(Trim(Request.form("reject")))
'		sRecschTime 	= UCase(Trim(Request.form("recschtime")))

		' 半角カンマチェック
		If InStr(sShipper,",")<>0 Or InStr(sShipLine,",")<>0 Or InStr(sVslCode,",")<>0 Or _
			InStr(sBL,",")<>0 Or InStr(sCont,",")<>0 _
		Then
			bError = true
			strError = "入力の際、半角カンマは使用しないで下さい。"
		Else

' トランザクションファイル作成
			' テンポラリファイル名を作成して、セッション変数に設定
			Dim sIM19, iSeqNo_IM19, strFileName, sTran, sTusin, sDate

			'' 更新時は前データを削除
			If bKind=0 And sDel="" Then
				'シーケンス番号
				iSeqNo_IM19 = GetDailyTransNo
				'通信日時取得
				sTusin  = SetTusinDate

				sIM19 = iSeqNo_IM19 & "," & sTranID & ",D,"  & sTusin & ",Web - " & sForwarder & ",," & _
				sShipper & "," & sShipLine & "," & sVslCode & "," &  sBLold & "," & sContold & "," & sForwarder
				sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM19
				strFileName_02 = "./send/" & sFileName & "." & SEND_EXTENT

				Set ti=fs.OpenTextFile(Server.MapPath(strFileName_02),2,True)
				ti.WriteLine sIM19
				ti.Close
				Set ti = Nothing
			End If

			'シーケンス番号
			iSeqNo_IM19 = GetDailyTransNo
			'通信日時取得
			sTusin  = SetTusinDate
			' 処理区分
			If sSend<>"" Then
				sSyori="R"
			Else
				sSyori="D"
			End If

			sIM19 = iSeqNo_IM19 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & sForwarder & ",," & _
			sShipper & "," & sShipLine & "," & sVslCode & "," &  sBL & "," & sCont & "," & sForwarder
			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM19
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT

			Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
			ti.WriteLine sIM19
			ti.Close
			Set ti = Nothing
	' トランザクションここまで


' Tempファイル作成
'			If sBL="" Then sBL="*"
'			If sCont="" Then sCont="*"
'			If sBLold="" Then sBLold="*"
'			If sContold="" Then sContold="*"

			' File System Object の生成
			Set fs=Server.CreateObject("Scripting.FileSystemobject")

			Dim strTempFileName
'			If bKind=1 Then
'				' テンポラリファイル名を作成して、セッション変数に設定
'				strTempFileName = GetNumStr(Session.SessionID, 8) & ".csv"
'				Session.Contents("tempfile")=strTempFileName
'			Else
				' 表示ファイルの取得
				strTempFileName = Session.Contents("tempfile")
				If strTempFileName="" Then
					' セッションが切れているとき
					Response.Redirect "sokuji-kaika-updtchk.asp"             'メニュー画面へ
					Response.End
				End If
'			End If

			strTempFileName="./temp/" & strTempFileName

			' 表示ファイルのOpen
			Set ti=fs.OpenTextFile(Server.MapPath(strTempFileName),1,True)

			' 詳細表示行のデータの取得
			Dim strData()
			LineNo=0
			Do While Not ti.AtEndOfStream
				strTemp=ti.ReadLine
				ReDim Preserve strData(LineNo)
				strData(LineNo) = strTemp
				LineNo=LineNo+1
			Loop
			ti.Close

			Set ti=fs.OpenTextFile(Server.MapPath(strTempFileName),2,True)

		' 更新時
			If bKind=0 Then

				ConnectSvr conn, rsd
				'' DBの読み込み
				sql = "SELECT NameAbrev FROM mShipper WHERE Shipper='" & sShipper & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sShipperAbrev = Trim(rsd("NameAbrev"))
				  rsd.MoveNext
				Loop
				rsd.Close
				'' DBの読み込み
				sql = "SELECT NameAbrev FROM mShipLine WHERE ShipLine='" & sShipLine & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sShipLineAbrev = Trim(rsd("NameAbrev"))
				  rsd.MoveNext
				Loop
				rsd.Close
				'' DBの読み込み
				sql = "SELECT FullName FROM mVessel WHERE VslCode='" & sVslCode & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sVesselFull = Trim(rsd("FullName"))
				  rsd.MoveNext
				Loop
				rsd.Close

				anyTmp=Split(strData(iLineNo-1),",")
				anyTmp(0) = sShipperAbrev
				anyTmp(1) = sShipLineAbrev
				anyTmp(2) = sVesselFull
'				If sBL<>"*" Then
					anyTmp(3) = sBL
'				ElseIf sCont<>"*" Then
					anyTmp(4) = sCont
'				Else
'					anyTmp(3) = ""
'				End If
				anyTmp(5) = sOpe
				anyTmp(6) = sOpeTel
				anyTmp(7) = ""
				anyTmp(8) = ""
'				anyTmp(6) = sReject
'				anyTmp(7) = sRecschTime
				anyTmp(9) = sShipper
				anyTmp(10) = sShipLine
				anyTmp(11) = sVslCode

				For i=1 To LineNo
					If i<>CInt(iLineNo) Then
					    ti.WriteLine strData(i-1)
					Else
						If sDel="" Then
							strTemp=anyTmp(0)
							For j=1 To UBound(anyTmp)
							    strTemp=strTemp & "," & anyTmp(j)
							Next
							ti.WriteLine strTemp
						End If
					End If
				Next
				ti.Close


			' 新規登録時
			Else
				Dim strTemp
'				If bKind=1 Then
					For i=1 To LineNo
						ti.WriteLine strData(i-1)
					Next
'				End If

				Dim sShipperAbrev,sShipLineAbrev,sVesselFull',sBlcont
'				If sBL<>"*" Then
'					sBlcont=sBL
'				Else
'					sBlcont=sCont
'				End If

				ConnectSvr conn, rsd
				'' DBの読み込み
				sql = "SELECT NameAbrev FROM mShipper WHERE Shipper='" & sShipper & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sShipperAbrev = Trim(rsd("NameAbrev"))
				  rsd.MoveNext
				Loop
				rsd.Close
				'' DBの読み込み
				sql = "SELECT NameAbrev FROM mShipLine WHERE ShipLine='" & sShipLine & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sShipLineAbrev = Trim(rsd("NameAbrev"))
				  rsd.MoveNext
				Loop
				rsd.Close
				'' DBの読み込み
				sql = "SELECT FullName FROM mVessel WHERE VslCode='" & sVslCode & "'"
				rsd.Open sql, conn, 0, 1, 1
				Do While Not rsd.EOF
				  sVesselFull = Trim(rsd("FullName"))
				  rsd.MoveNext
				Loop
				rsd.Close

				strTemp = sShipperAbrev & "," &  sShipLineAbrev & "," & sVesselFull & "," & sBL & "," & sCont & "," & _
									sOpe & "," & sOpeTel & "," & "" & "," & "" & "," & _
									sShipper & "," & sShipLine & "," & sVslCode

				ti.WriteLine strTemp
				ti.Close
			End If

' Tempここまで

			If sDel<>"" Or bKind=0 Then
				If sDel<>"" Then
					WriteLog fs, "7002", "即時搬出システム-海貨用申込み", "13", sShipper & "/" & sShipLine & "/" & sVslCode & "/" & sBL & "/" & sCont & ","
				Else
					WriteLog fs, "7002", "即時搬出システム-海貨用申込み", "12", sShipper & "/" & sShipLine & "/" & sVslCode & "/" & sBL & "/" & sCont & ",入力内容の正誤:0(正しい)"
				End If
				Response.Redirect "sokuji-kaika-list.asp"
				Response.End
			End If

			' エラーメッセージの表示
			If bKind<>0 Then
				strError = "正常に送信されました。"
				Session.Contents("kind") = 2
			End If


		End If
	End If

'		If sBL="*" Then sBL=""
'		If sCont="*" Then sCont=""

%>

<html>
<head>
<title>即時搬出申込み（海貨）</title>
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
          <td rowspan=2><img src="gif/sokuji1t.gif" width="506" height="73"></td>
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
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>（海貨用）即時搬出申込み</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>
			<BR>
              <center>

              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" width=70 height=23 align=center valign=middle>
                    <font color="#FFFFFF"><b>荷主コード</b></font>
                  </td>
                  <td width=100>
					&nbsp;<%=sShipper%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" width=70 height=23 align=center valign=middle>
                    <font color="#FFFFFF"><b>船社コード</b></font>
                  </td>
                  <td nowrap>
					&nbsp;<%=sShipLine%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" width=70 height=23 align=center valign=middle>
                    <font color="#FFFFFF"><b>船名コード</b></font>
                  </td>
                  <td nowrap>
					&nbsp;<%=sVslCode%>
                  </td>
                </tr>
<%
	Dim sBLContName,sBLCont
	If sBL="" Then
		sBLContName = "コンテナNo."
		sBLCont     = sCont
	Else
		sBLContName = "BL No."
		sBLCont     = sBL
	End If
%>
                <tr> 
                  <td bgcolor="#000099" width=70 height=23 align=center valign=middle>
                    <font color="#FFFFFF"><b><%=sBLContName%></b></font>
                  </td>
                  <td nowrap>
					&nbsp;<%=sBLCont%>
                  </td>
                </tr>

			  </table>
				<BR>

<%
	If bError Then
		DispErrorMessage strError

	ElseIf bKind<>0 Then
		DispInformationMessage strError

	End If
	
	
%>
<BR>
<form action="JavaScript:window.history.back()" id=form1 name=form1>
	<input type=button value=" 戻  る " onclick="history.back()">
</form>

              </center>
		  </td>
		</tr>
	  </table>

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
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

<%
	' Log作成
	If bError Then
		If bKind<>0 Then
			WriteLog fs, "7002", "即時搬出システム-海貨用申込み", "11", sShipper & "/" & sShipLine & "/" & sVslCode & "/" & sBL & "/" & sCont & ",入力内容の正誤:1(誤り)"
		Else
			WriteLog fs, "7002", "即時搬出システム-海貨用申込み", "12", sShipper & "/" & sShipLine & "/" & sVslCode & "/" & sBL & "/" & sCont & ",入力内容の正誤:1(誤り)"
		End If

	ElseIf bKind<>0 Then
		WriteLog fs, "7002", "即時搬出システム-海貨用申込み", "11", sShipper & "/" & sShipLine & "/" & sVslCode & "/" & sBL & "/" & sCont & ",入力内容の正誤:0(正しい)"
	Else
		WriteLog fs, "7002", "即時搬出システム-海貨用申込み", "12", sShipper & "/" & sShipLine & "/" & sVslCode & "/" & sBL & "/" & sCont & ",入力内容の正誤:0(正しい)"
	End If
%>
