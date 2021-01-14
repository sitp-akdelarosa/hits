<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	【海貨入力】	エラーチェック、表示、ファイル作成
%>

<%
    ' セッションのチェック
    CheckLogin "pickselect.asp"

	Dim bKind,sSend,sStop,sDel,iLineNo
	' 新規(1) or 更新(0)
    bKind = Trim(Session.Contents("kind"))
	' 種別
    sSend 	= Trim(Request.form("send"))
    sStop 	= Trim(Request.form("stop"))
    sDel 	= Trim(Request.form("del"))
    iLineNo	= Trim(Request.form("lineno"))

    strUserKind=Session.Contents("userkind")

	If bKind=1 And sStop<>"" Then
        Response.Redirect "pickexp-dblist.asp"

	ElseIf bKind=2 And sStop<>"" Then
        Response.Redirect "pickexp-list.asp"

	Else
		' トランザクションファイルの拡張子 
		Const SEND_EXTENT = "snd"
		' トランザクションＩＤ
		Const sTranID = "EX16"
		' 送信場所
		Const sPlace = ""
	    ' セッションのチェック
	    CheckLogin "pickselect.asp"
	    ' エラーフラグのクリア
	    bError = false
	    ' 入力フラグのクリア
	    bInput = true
	    ' 指定引数の取得
	    Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sTraderCode,iSize,sType,iHeight,sPick,sEmpDate
		Dim sEmpDateS,sShipLine,sRecDate,sRecDateS
	    sSosin 	= UCase(Trim(Request.form("forwarder")))
	    sUser 	= UCase(Trim(Request.form("user")))
	    sUserNo = UCase(Trim(Request.form("userno")))
	    sVslCode = UCase(Trim(Request.form("vslcode")))
	    sVoyCtrl = UCase(Trim(Request.form("voyctrl")))
	    sBooking = UCase(Trim(Request.form("booking")))
	    sTraderCode = UCase(Trim(Request.form("tradercode")))
	    iSize 	= Trim(Request.form("size"))
	    sType 	= UCase(Trim(Request.form("type")))
	    iHeight = Trim(Request.form("height"))
	    sPick 	= UCase(Trim(Request.form("pickplace")))
	    sRemark = UCase(Trim(Request.form("remark")))
	    sOpeCode = UCase(Trim(Request.form("opecode")))
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

		sRecDate = Trim(Request.form("recdate_year")) 
		sRecDate = sRecDate & Right("0" & Trim(Request.form("recdate_mon")),2)
		sRecDate = sRecDate & Right("0" & Trim(Request.form("recdate_day")),2)
		If sRecDate=00 Then
			sRecDate = ""
		Else
			sRecDateT = Trim(Request.form("recdate_year")) 
			sRecDateT = sRecDateT & "/" & Right("0" & Trim(Request.form("recdate_mon")),2)
			sRecDateT = sRecDateT & "/" & Right("0" & Trim(Request.form("recdate_day")),2)
			sRecDateS = Trim(Request.form("recdate_year")) 
			sRecDateS = sRecDateS & "年 " & Trim(Request.form("recdate_mon"))
			sRecDateS = sRecDateS & "月 " & Trim(Request.form("recdate_day")) & "日 "
		End If

		If strUserKind="荷主" Then
			sTraderCode = ""
			sOpeCode = ""
			sRecDate = ""
			sRecDateT = ""
			sRecDateS = ""
		End If


	    ' File System Object の生成
	    Set fs=Server.CreateObject("Scripting.FileSystemobject")

		' 半角カンマチェック
		If InStr(sVslCode,",")<>0 Or _
			InStr(sVoyCtrl,",")<>0 Or _
			InStr(sBooking,",")<>0 Or _
			InStr(sTraderCode,",")<>0 Or _
			InStr(sRemark,",")<>0 Or _
			InStr(sPick,",")<>0 Or _
			InStr(sUser,",")<>0 Or _
			InStr(sUserNo,",")<>0 _
		Then

		    bError = true
			strError = "入力の際、半角カンマは使用しないで下さい。"

		Else

			ConnectSvr conn, rsd
			' 荷主コードと荷主管理番号の重複チェック
			Dim iRecCount
			If Not bKind=0 Then
				sql = "SELECT count(*) FROM ExportCargoInfo WHERE Shipper='" & sUser & "' AND ShipCtrl='" & sUserNo & "'"
				rsd.Open sql, conn, 0, 1, 1
				If Not rsd.EOF Then
					iRecCount = rsd(0)
					If Not iRecCount=0 Then
					    bError = true
						strError = "荷主コードと荷主管理番号が重複しています。"
					End If
				End If
				rsd.Close
			End If

'項目についてはチェックしないように変更	2002/3/8

			' 船名が存在するか
'			sql = "SELECT count(*) FROM VslSchedule WHERE VslCode='" & sVslCode & "' AND LdVoyage='" & sVoyCtrl & "'" 
'			rsd.Open sql, conn, 0, 1, 1
'			If Not rsd.EOF Then
'				iRecCount = rsd(0)
'				If iRecCount=0 Then
'				    bError = true
'					strError = "船名（コールサイン）とVoyage No.が一致しません。"
'				End If
'			End If
'			rsd.Close

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
	    Dim sEX16, iSeqNo_EX16, strFileName, sTran, sTusin, sDate
		'シーケンス番号
		iSeqNo_EX16 = GetDailyTransNo
		'通信日時取得
		sTusin  = SetTusinDate

		If strUserKind="海貨" Then
			sSender = sSosin
		Else
			sSender = sUser
		End If

		sEX16 = iSeqNo_EX16 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSender & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sBooking & "," & sUser & "," & sUserNo & "," & sSosin & "," & _
				sContainer & "," & iSize & "," & sType & "," & iHeight & "," & sRemark & "," & sTraderCode & "," & _
				sEmpDate & ",," & sPick & "," & sRecDate & "," & sOpeCode
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX16
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sEX16
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
	            anyTmp(3) = sUserNo
	            anyTmp(4) = sBooking
	            anyTmp(5) = sTraderCode
	            anyTmp(6) = sEmpDateT
	            anyTmp(7) = ""
	            anyTmp(8) = iSize
	            anyTmp(9) = sType
	            anyTmp(10) = iHeight
	            anyTmp(11) = sRemark
	            anyTmp(12) = sPick
	            anyTmp(13) = sRecDateT
	            anyTmp(14) = sOpeCode
				anyTmp(15) = sSosin

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

				strTemp = sVslCode & "," &  sVoyCtrl & "," & sUser & "," & sUserNo & "," & _
						 sBooking & "," & sTraderCode & "," & sEmpDateT & ",," & _
						 iSize & "," & sType & "," & iHeight & "," & sRemark & "," &_
						 sPick & "," & sRecDateT & "," & sOpeCode & "," & sSosin

                ti.WriteLine strTemp
	            ti.Close

			End If

		End If

' Tempここまで

	' ログファイル書き出し
	Dim sRLogDate,sLogDate,sLogTime
	sRLogDate = Trim(Request.form("recdate_year")) & "/"
	sRLogDate = sRLogDate & Right("0" & Trim(Request.form("recdate_mon")),2) & "/"
	sRLogDate = sRLogDate & Right("0" & Trim(Request.form("recdate_day")),2)
	sLogTime = Trim(Request.form("emparvtime_year")) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_mon")),2) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_day")),2) & " "
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_hour")),2) & ":"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_min")),2)
	If sEmpDateT="" Then
		sLogTime = ""
	End If
	If sRecDateT="" Then
		sRLogDate = ""
	End If

	strOption = sVslCode & _
				"/" & sVoyCtrl & _
				"/" & sSosin & _
				"/" & sUser & _
				"/" & sUserNo & _
				"/" & sBooking
	If strUserKind="海貨" Then
				strOption = strOption & "/" & sOpeCode & "/" & sTraderCode
	End If
	strOption = strOption & "/" & sLogTime & _
				"/" & iSize & _
				"/" & sType & _
				"/" & iHeight & _
				"/" & sPick
	If strUserKind="海貨" Then
				strOption = strOption & "/" & sRLogDate
	End If
	strOption = strOption & "/" & sRemark

    If bError Then
		strOption = strOption &	",入力内容の正誤:1(誤り)"
    Else
		strOption = strOption & ",入力内容の正誤:0(正しい)"
    End If

	If strUserKind="海貨" Then
		iNum = "a111"
	Else
		iNum = "a114"
	End If

	If bKind=1 Then
		'新規
   		WriteLog fs, iNum,"空コンピックアップシステム-" & strUserKind & "用依頼入力", "11", strOption
	ElseIf sDel<>"" Then
   		WriteLog fs, iNum,"空コンピックアップシステム-" & strUserKind & "用依頼入力", "13", strOption
	Else
   		WriteLog fs, iNum,"空コンピックアップシステム-" & strUserKind & "用依頼入力", "12", strOption
	End If

    If Not bError And bKind=0 Then
		Response.Redirect "pickexp-list.asp"
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
<%
	If strUserKind="海貨" Then
		titlegif = "pickkat"
	Else
		titlegif = "picknit"
	End If
%>
          <td rowspan=2><img src="gif/<%=titlegif%>.gif" width="506" height="73"></td>
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
                    <font color="#FFFFFF"><b>船名</b></font>
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
                  <td bgcolor="#000099" nowrap align=center valign=middle>
					<font color="#FFFFFF"><b>荷主コード</b></font>
				  </td>
                  <td nowrap>
					<%=sUser%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>荷主管理番号</b></font>
                  </td>
                  <td nowrap>
					<%=sUserNo%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Booking No.</b></font>
                  </td>
                  <td nowrap>
					<%=sBooking%>
                  </td>
                </tr>

<% 	If strUserKind="海貨" Then %>
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>港運コード</b></font>
                  </td>
                  <td nowrap>
					<% If sOpeCode<>"" Then %>
						<%=sOpeCode%>
					<% Else %>
						<BR>
					<% End If %>
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
<% End If %>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>空コン倉庫到着指定日時</b></font>
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
                    <font color="#FFFFFF"><b>高さ</b></font>
                  </td>
                  <td nowrap>
					<% If iHeight<>"" Then %>
						<%=iHeight%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>空コンピック場所</b></font>
                  </td>
                  <td nowrap>
					<% If sPick<>"" Then %>
						<%=sPick%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

<% 	If strUserKind="海貨" Then %>
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>空コン搬出指定日</b></font>
                  </td>
                  <td nowrap>
					<% If sRecDate<>"" Then %>
						<%=sRecDateS%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>
<% End If %>

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
<form>
	<input type=button value=" 戻  る " onClick="JavaScript:window.history.back()">
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
