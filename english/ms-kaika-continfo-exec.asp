<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	【コンテナ情報入力】	エラーチェック、表示、ファイル作成
	'
%>

<%
    ' セッションのチェック
    CheckLogin "ms-kaika.asp"

	Dim bKind,iLineNo
	' 新規(1) or 更新(0)
    bKind = Trim(Request.form("kind"))
    iLineNo	= Trim(Request.form("lineno"))

	' トランザクションファイルの拡張子 
	Const SEND_EXTENT = "snd"
	' トランザクションＩＤ
	Const sTranID05 = "EX05"
	Const sTranID16 = "EX16"
	' 処理区分
	Const sSyori = "R"
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
	Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sCont,sSeal,sCargoWeight,sContWeight,sRifer,sDanger
    sUser 	= UCase(Trim(Request.form("user")))
    sUserNo = UCase(Trim(Request.form("userno")))
    sVslCode = UCase(Trim(Request.form("vslcode")))
    sVoyCtrl = UCase(Trim(Request.form("voyctrl")))
    sBooking = UCase(Trim(Request.form("booking")))
	sCont 		= UCase(Trim(Request.form("cont")))
	sSeal 		= UCase(Trim(Request.form("seal")))
	sCargoWeight= UCase(Trim(Request.form("cargow")))
	sContWeight	= UCase(Trim(Request.form("contw")))
	sRifer 		= UCase(Trim(Request.form("rifer")))
	sDanger 	= UCase(Trim(Request.form("danger")))
	iLineNo		= Request.form("lineno")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ConnectSvr conn, rsd
	' コンテナNo.が存在するか
	sql = "SELECT ExportCont.VanTime,VslSchedule.ShipLine " & _
		  "FROM ExportCont,VslSchedule " & _
		  "WHERE " & _
			"ExportCont.VslCode='" & sVslCode & "' AND " & _
			"ExportCont.ContNo='" & sCont & "' AND " & _
			"ExportCont.BookNo='" & sBooking & "' AND " & _
			"VslSchedule.VslCode='" & sVslCode & "'"
	rsd.Open sql, conn, 0, 1, 1

	Dim sVanTime,sShipLine
	If Not rsd.EOF Then
	    sVanTime  = Trim(rsd("VanTime"))
	    sShipLine = Trim(rsd("ShipLine"))
	Else
	    bError = true
		strError = "指定されたコンテナNo.が存在しません。"
	End If
	rsd.Close


    If Not bError Then

' トランザクションファイル作成

	    ' テンポラリファイル名を作成して、セッション変数に設定
	    Dim sEX16, sEX05, iSeqNo_EX16, iSeqNo_EX05, strFileName, sTran, sTusin, sDate

		'通信日時取得
		sTusin  = SetTusinDate

		If bKind=1 Then
			' EX16用パラメータの取得
			Dim sContSize,sContType,sContHeight,sRemark,sTrucker,sWHArTime,sCYRecDate,sPickPlace
			sql = "SELECT ContSize,ContType,ContHeight,Remark,Trucker,WHArTime,CYRecDate,PickPlace " & _
				  "FROM ExportCargoInfo " & _
				  "WHERE " & _
					"Shipper='" & sUser & "' AND " & _
					"ShipCtrl='" & sUserNo & "'"
			rsd.Open sql, conn, 0, 1, 1

			If Not rsd.EOF Then
			    sContSize 	= Trim(rsd("ContSize"))
			    sContType 	= Trim(rsd("ContType"))
			    sContHeight = Trim(rsd("ContHeight"))
			    sRemark 	= Trim(rsd("Remark"))
			    sTrucker 	= Trim(rsd("Trucker"))
			    sWHArTime 	= Trim(rsd("WHArTime"))
			    sCYRecDate 	= Trim(rsd("CYRecDate"))
			    sPickPlace 	= Trim(rsd("PickPlace"))
			Else
			    bError = true
				strError = "指定されたコンテナNo.が存在しません。"
			End If
			rsd.Close

			'シーケンス番号
			iSeqNo_EX16 = GetDailyTransNo
			'倉庫到着指定日時
			If sWHArTime<>"" Then
				sWHArTime = "20" & Left(sWHArTime,2) & Mid(sWHArTime,4,2) & Mid(sWHArTime,7,2) & _
							Mid(sWHArTime,10,2) & Mid(sWHArTime,13,2)
			End If
			'ＣＹ搬入指定日
			If sCYRecDate<>"" Then
				sCYRecDate = "20" & Left(sCYRecDate,2) & Mid(sCYRecDate,4,2) & Mid(sCYRecDate,7,2)
			End If

			sEX16 = iSeqNo_EX16 & "," & sTranID16 & "," & sSyori & ","  & sTusin & ",Web - " & _
					sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
					sBooking & "," & sUser & "," & sUserNo & "," & sSosin & "," & _
					sCont & "," & sContSize & "," & sContType & "," & sContHeight & "," & _
					sRemark & "," & sTrucker & "," & _
					sWHArTime & "," & sCYRecDate & "," & sPickPlace
			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX16
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
			ti.WriteLine sEX16
		    ti.Close
			Set ti = Nothing

		End If

		'シーケンス番号
		iSeqNo_EX05 = GetDailyTransNo
		'バン詰め日時
		If sVanTime<>"" Then
			sVanTime = "20" & Left(sVanTime,2) & Mid(sVanTime,4,2) & Mid(sVanTime,7,2) & _
						Mid(sVanTime,10,2) & Mid(sVanTime,13,2)
		End If

		sEX05 = iSeqNo_EX05 & "," & sTranID05 & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sCont & "," & sBooking & "," & sShipLine & "," & sVanTime & "," & _
				sContWeight & "," & sSeal & "," & sCargoWeight & "," & sSosin & ",," & _
				sRifer & sDanger
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX05
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sEX05
	    ti.Close
		Set ti = Nothing

' トランザクションここまで


' Tempファイル作成

	    ' File System Object の生成
	    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	    Dim strTempFileName
	    ' 表示ファイルの取得
	    strTempFileName = Session.Contents("tempfile")
	    If strTempFileName="" Then
	        ' セッションが切れているとき
	        Response.Redirect "http://www.hits-h.com/index.asp"             'メニュー画面へ
	        Response.End
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

  		anyTmp=Split(strData(iLineNo-1),",")
        anyTmp(0) = sVslCode
        anyTmp(1) = sVoyCtrl
        anyTmp(2) = sUser
        anyTmp(3) = sUserNo
        anyTmp(4) = sBooking
        anyTmp(5) = sCont
        anyTmp(6) = sSeal
        anyTmp(7) = sCargoWeight
        anyTmp(8) = sContWeight
        anyTmp(9) = sRifer
        anyTmp(10) = sDanger

        For i=1 To LineNo
            If i<>CInt(iLineNo) Then
                ti.WriteLine strData(i-1)
            Else
                strTemp=anyTmp(0)
                For j=1 To UBound(anyTmp)
                    strTemp=strTemp & "," & anyTmp(j)
                Next
                ti.WriteLine strTemp
            End If
        Next
        ti.Close

' Tempここまで

		Response.Redirect "ms-kaika-continfo-list.asp"

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
          <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
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
<%
        DispErrorMessage strError

	    strOption = "BL No," & sBLNo & ",コンテナNo.," & sContNo & ",時刻," & sDate & "," & "入力内容の正誤:0(正しい)"

        ' 海貨入力項目選択
        WriteLog fs, "輸出貨物情報入力", strOption
%>

			<form>
				<input type=button value=" 戻  る " onClick="JavaScript:window.history.back()">
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
<!-------------ログイン画面終わり--------------------------->
<%
    DispMenuBarBack "ms-kaika-continfo.asp"
%>
</body>
</html>

<%
%>
