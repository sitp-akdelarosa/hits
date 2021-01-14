<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
'通信日時のセット
Function SetTusinDate()
	'戻り値		[ O ]通信日時(文字列)
				'YYYYMMDDHHNNSS
	SetTusinDate = Trim(Year(Date)) & _
				   Trim(Right("0" & Month(Date), 2)) & _
				   Trim(Right("0" & Day(Date), 2)) & _
				   Trim(Right("0" & Hour(Time), 2)) & _
				   Trim(Right("0" & Minute(Time), 2)) & _
				   Trim(Right("0" & Second(Time), 2))
End Function

' 数値を指定した桁数の文字列に変換(右詰・余白には０)コンテナ入力用
Function ArrangeNumV(nNumber, nFigure)
	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNumV = sNum
End Function

    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "index.asp"

    ' 指定引数の取得
    Dim iLine         '入力行
    iLine=Session.Contents("lineary")
	Session.Contents("lineary") = ""

	sLoginKind = Session.Contents("userkind")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"             'メニュー画面へ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

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

    ' トランザクションファイルの拡張子 
    Const SEND_EXTENT = "snd"
    ' トランザクションＩＤ
    sTranID = "EX18"
    ' 処理区分
    Const sSyori = "R"
    ' 送信場所
    Const sPlace = ""

    ' セッションのチェック
    CheckLogin "expentry.asp"
    sSosin = Trim(Session.Contents("userid"))

	' タイトル取得
	strTitle = Trim(Request.form("title"))


    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim sEXxx, iSeqNo, strFileName_01, sTran, sTusin, sDate, sPickPlace

    'シーケンス番号
    iSeqNo = GetDailyTransNo

    '通信日時取得
    sTusin  = SetTusinDate
    sPickPlace = Trim(Request.form("pickplace")) 
    sDate = Trim(Request.form("pickyear")) 
    sDate = sDate & Right("0" & Trim(Request.form("pickmon")),2)
    sDate = sDate & Right("0" & Trim(Request.form("pickday")),2)
	If sDate="00" Then
		sDate = ""
	End If
    strTemp=Left(sDate,4) & "/" & Mid(sDate,5,2) & "/" & Mid(sDate,7,2)

    sLogDate = Trim(Request.form("pickyear")) & "/"
    sLogDate = sLogDate & Right("0" & Trim(Request.form("pickmon")),2) & "/"
    sLogDate = sLogDate & Right("0" & Trim(Request.form("pickday")),2)
	If sLogDate="/0/0" Then
		sLogDate=""
	End If

	If sLoginKind="港運" Then
	    strOption = sPickPlace & "/" & sLogDate & "," & "入力内容の正誤:0(正しい)"
	Else
	    strOption = sLogDate & "," & "入力内容の正誤:0(正しい)"
	End If

    sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
    strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

	iLineAry = Split(iLine,",")
    'トランザクションファイル作成
	For k=0 To UBound(iLineAry)
	    ' エラーフラグのクリア
	    bError = false

	    If Not bError Then
		    'シーケンス番号
		    iSeqNo = GetDailyTransNo

	        anyTmp=Split(strData(iLineAry(k)-1),",")
	        sEXxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
	                sSosin & "," & sPlace & "," & anyTmp(25) & "," & anyTmp(4) & "," & anyTmp(3) & "," &_
					anyTmp(0) & "," & anyTmp(23) & "," & anyTmp(14) & ",2," &_
					sPickPlace & "," & sDate

	        ti.WriteLine sEXxx
	    End If
	Next

    ti.Close

    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    ' テンポラリファイル更新
    For i=1 To LineNo
		bSameFlag = false
		For l=0 To UBound(iLineAry)
			If i=CInt(iLineAry(l)) Then
				bSameFlag = true
			End If
		Next

        If Not bSameFlag Then
            ti.WriteLine strData(i-1)
        Else
	        anyTmp=Split(strData(i-1),",")
            strTemp=anyTmp(0)
			If sPickPlace<>"" Then
				anyTmp(20) = sPickPlace
				anyTmp(27) = "1"
				anyTmp(26) = "2"
			End If
			If sLogDate<>"" Then
			    anyTmp(24) = sLogDate
				anyTmp(28) = "1"
				anyTmp(26) = "2"
			End If
            For j=1 To UBound(anyTmp)
                strTemp=strTemp & "," & anyTmp(j)
            Next
            ti.WriteLine strTemp
        End If
    Next

    ti.Close

    ' 海貨入力項目選択
	If sLoginKind="港運" Then
	    WriteLog fs, "a109","空コンピックアップシステム-空コン受取場所・搬出日入力","12", strOption
	    Response.Redirect "picklist.asp?kind=4"
	Else
	    WriteLog fs, "a109","空コンピックアップシステム-空コン受取場所・搬出日入力","11", strOption
		Response.Redirect "picklist.asp?kind=2"
	End If

    Response.End

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
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/expkoun.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strRoute = Session.Contents("route")
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
				<%=strRoute%> &gt; 空コン受取場所・搬出日入力
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
          <td nowrap><b>
<%
    Response.Write strTitle
%>
            入力</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
	 	<BR>

<%
    ' エラーメッセージの表示
    DispErrorMessage strError 
    strOption = sPickPlace & "/" & sLogDate & "," & "入力内容の正誤:1(誤り)"
    ' 海貨入力項目選択
    WriteLog fs, "a109","空コンピックアップシステム-空コン受取場所・搬出日入力","10", strOption
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
<!-------------登録画面終わり--------------------------->
<%
    strTemp = "picklist.asp?kind=" & iLoginKind
    DispMenuBarBack strTemp
%>
</body>
</html>
<%

%>