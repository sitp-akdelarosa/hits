<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	【輸入コンテナ情報入力】	エラーチェック、表示、ファイル作成
	'
%>

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"

	Dim bKind,sSend,sStop,sDel,iLineNo
	' 新規(1) or 更新(0)
    bKind = Trim(Session.Contents("kind"))
	' 種別
    sSend 	= Trim(Request.form("send"))
    sStop 	= Trim(Request.form("stop"))
    sDel 	= Trim(Request.form("del"))
    iLineNo	= Trim(Request.form("lineno"))

	If bKind=1 And sStop<>"" Then
        Response.Redirect "ms-kaika-impcontinfo-updatecheck.asp"

	ElseIf bKind=2 And sStop<>"" Then
        Response.Redirect "ms-kaika-impcontinfo-list.asp"

	Else
		' トランザクションファイルの拡張子 
		Const SEND_EXTENT = "snd"
		' トランザクションＩＤ
		Const sTranID = "IM18"
		' 送信場所
		Const sPlace = ""
	    ' セッションのチェック
	    CheckLogin "ms-kaika.asp"
		sSosin = Trim(Session.Contents("userid"))
	    ' エラーフラグのクリア
	    bError = false
	    ' 入力フラグのクリア
	    bInput = true
	    ' 指定引数の取得
	    Dim sVslCode,sVoyCtrl,sUser,sCont,sBL,sTraderCode,iSize,sType,sRemark,sEmpDate,sEmpDateS
	    sUser 	= UCase(Trim(Request.form("user")))
	    sCont 	= UCase(Trim(Request.form("cont")))
	    sVslCode = UCase(Trim(Request.form("vslcode")))
	    sVoyCtrl = UCase(Trim(Request.form("voyctrl")))
	    sBL		 = UCase(Trim(Request.form("bl")))
	    sTraderCode = UCase(Trim(Request.form("tradercode")))
	    iSize 	= Trim(Request.form("size"))
	    sType 	= UCase(Trim(Request.form("type")))
	    sRemark = UCase(Trim(Request.form("remark")))
		sEmpDate = Trim(Request.form("emparvtime_year")) 
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_mon")),2)
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_day")),2)
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_hour")),2)
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_min")),2)
		If sEmpDate=0000 Then
			sEmpDate = ""
		Else
			sEmpDateT = Trim(Request.form("emparvtime_year")) 
			sEmpDateT = sEmpDateT & "/" & Right("0" & Trim(Request.form("emparvtime_mon")),2)
			sEmpDateT = sEmpDateT & "/" & Right("0" & Trim(Request.form("emparvtime_day")),2)
			sEmpDateT = sEmpDateT & " " & Right("0" & Trim(Request.form("emparvtime_hour")),2)
			sEmpDateT = sEmpDateT & ":" & Right("0" & Trim(Request.form("emparvtime_min")),2)
			sEmpDateS = Trim(Request.form("emparvtime_year")) 
			sEmpDateS = sEmpDateS & "年 " & Trim(Request.form("emparvtime_mon"))
			sEmpDateS = sEmpDateS & "月 " & Trim(Request.form("emparvtime_day"))
			sEmpDateS = sEmpDateS & "日 " & Trim(Request.form("emparvtime_hour"))
			sEmpDateS = sEmpDateS & "時 " & Trim(Request.form("emparvtime_min")) & "分 "
		End If


	    ' File System Object の生成
	    Set fs=Server.CreateObject("Scripting.FileSystemobject")

		' 半角カンマチェック
		If InStr(sVslCode,",")<>0 Or _
			InStr(sVoyCtrl,",")<>0 Or _
			InStr(sBL,",")<>0 Or _
			InStr(sTraderCode,",")<>0 Or _
			InStr(sRemark,",")<>0 Or _
			InStr(sCont,",")<>0 Or _
			InStr(sUser,",")<>0 _
		Then

		    bError = true
			strError = "入力の際、半角カンマは使用しないで下さい。"

		Else

			ConnectSvr conn, rsd
			' 船名と次航とコンテナNo.の重複チェック
			Dim iRecCount
			If Not bKind=0 Then
				sql = "SELECT count(*) FROM ImportCargoInfo " & _
						"WHERE VslCode='" & sVslCode & "' AND DsVoyage='" & sVoyCtrl & "' AND ContNo='" & sCont & "'"
				rsd.Open sql, conn, 0, 1, 1
				If Not rsd.EOF Then
					iRecCount = rsd(0)
					If Not iRecCount=0 Then
					    bError = true
						strError = "船名, Voyage No, コンテナNo.が重複しています。"
					End If
				End If
				rsd.Close
			End If

			' 船名が存在するか
			sql = "SELECT count(*) FROM mVessel WHERE VslCode='" & sVslCode & "'"
			rsd.Open sql, conn, 0, 1, 1
			If Not rsd.EOF Then
				iRecCount = rsd(0)
				If iRecCount=0 Then
				    bError = true
					strError = "指定された船名が存在しません。"
				End If
			End If
			rsd.Close

		End If
	End If

    If Not bError Then
		' 処理区分
		Dim sSyori
		If sSend<>"" Then
			sSyori = "R"
		Else
			sSyori = "D"
		End If

		Const sContainer = ""

' トランザクションファイル作成

	    ' テンポラリファイル名を作成して、セッション変数に設定
	    Dim sIM18, iSeqNo_IM18, strFileName, sTran, sTusin, sDate
		'シーケンス番号
		iSeqNo_IM18 = GetDailyTransNo
		'通信日時取得
		sTusin  = SetTusinDate

		sIM18 = iSeqNo_IM18 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sBL & "," & sUser & "," &  sSosin & "," & _
				sCont & "," & iSize & "," & sType & "," & sTraderCode & "," & _
				sRemark & "," & sEmpDate
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM18
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sIM18
	    ti.Close
		Set ti = Nothing

' トランザクションここまで


' Tempファイル作成

		    ' File System Object の生成
		    Set fs=Server.CreateObject("Scripting.FileSystemobject")

		    Dim strTempFileName
			If bKind=1 Then
			    ' テンポラリファイル名を作成して、セッション変数に設定
			    strTempFileName = GetNumStr(Session.SessionID, 8) & ".csv"
			    Session.Contents("tempfile")=strTempFileName

			Else
			    ' 表示ファイルの取得
			    strTempFileName = Session.Contents("tempfile")
			    If strTempFileName="" Then
			        ' セッションが切れているとき
			        Response.Redirect "http://www.hits-h.com/index.asp"             'メニュー画面へ
			        Response.End
			    End If

			End If

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

	      		anyTmp=Split(strData(iLineNo-1),",")
	            anyTmp(0) = sVslCode
	            anyTmp(1) = sVoyCtrl
	            anyTmp(2) = sUser
	            anyTmp(3) = sBL
	            anyTmp(4) = sCont
	            anyTmp(5) = sTraderCode
	            anyTmp(6) = sEmpDateT
	            anyTmp(7) = iSize
	            anyTmp(8) = sType
	            anyTmp(9) = sRemark

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

				If bKind=2 Then
		            For i=1 To LineNo
						ti.WriteLine strData(i-1)
					Next
				End If

				strTemp = sVslCode & "," &  sVoyCtrl & "," & sUser & "," & sBL & "," & _
						 sCont & "," & sTraderCode & "," & sEmpDateT & "," & _
						 iSize & "," & sType & "," & sRemark

                ti.WriteLine strTemp
	            ti.Close

			End If

		End If

' Tempここまで

	' ログファイル書き出し
	Dim sLogTime
	sLogTime = Trim(Request.form("emparvtime_year")) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_mon")),2) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_day")),2) & " "
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_hour")),2) & ":"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_min")),2)
	If sEmpDateT="" Then
		sLogTime = ""
	End If

	strOption = sVslCode & _
				"/" & sVoyCtrl & _
				"/" & sUser & _
				"/" & sBL & _
				"/" & sCont & _
				"/" & sTraderCode & _
				"/" & sLogTime & _
				"/" & iSize & _
				"/" & sType & _
				"/" & sRemark & ","

    If bError Then
		strOption = strOption &	"入力内容の正誤:1(誤り)"
    Else
		strOption = strOption & "入力内容の正誤:0(正しい)"
    End If

	If bKind=1 Then
  		WriteLog fs, "4110","海貨入力輸入コンテナ情報-情報入力","11", strOption
	ElseIf sDel<>"" Then
  		WriteLog fs, "4110","海貨入力輸入コンテナ情報-情報入力","13", strOption
	Else
  		WriteLog fs, "4110","海貨入力輸入コンテナ情報-情報入力","12", strOption
	End If


    If Not bError And bKind=0 Then
		Response.Redirect "ms-kaika-impcontinfo-list.asp"
	End If
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
<% If bKind<>0 Then %>
          <td nowrap><b>新規情報入力</b></td>
<% Else %>
          <td nowrap><b>更新情報入力</b></td>
<% End If %>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
<%
	If Not bError Then
%>

              <table border="1" cellspacing="2" cellpadding="3">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>船名(コールサイン)</b></font>
                  </td>
                  <td nowrap>
					<%=sVslCode%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Voyage No.</b></font>
                  </td>
                  <td nowrap>
					<%=sVoyCtrl%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>海貨コード</b></font>
                  </td>
                  <td nowrap>
					<%=sSosin%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle> <font color="#FFFFFF"><b>荷主コード</b></font></td>
                  <td nowrap>
					<%=sUser%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>BL No.</b></font>
                  </td>
                  <td nowrap>
					<%=sBL%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>コンテナNo.</b></font>
                  </td>
                  <td nowrap>
					<%=sCont%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>指定陸運業者コード</b></font>
                  </td>
                  <td nowrap>
					<% If sTraderCode<>"" Then %>
						<%=sTraderCode%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>実入り倉庫到着指定日時</b></font>
                  </td>
                  <td nowrap>
					<% If sEmpDate<>"" Then %>
						<%=sEmpDateS%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>サイズ</b></font>
                  </td>
                  <td nowrap>
					<% If iSize<>"" Then %>
						<%=iSize%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>タイプ</b></font>
                  </td>
                  <td nowrap>
					<% If sType<>"" Then %>
						<%=sType%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>倉庫略称</b></font>
                  </td>
                  <td nowrap>
					<% If sRemark<>"" Then %>
						<%=sRemark%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

              </table><BR>
<%
	    ' エラーメッセージの表示
		strError = "正常に送信されました。"

		Session.Contents("kind") = 2

	End If

		If bError Then
%><BR><%
	        DispErrorMessage strError
		Else
	        DispInformationMessage strError
%>
<BR>
<form action="JavaScript:window.history.back()">
	<input type=submit value=" 戻  る ">
</form>
<%
		End If
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
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

<%
%>
