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
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' 指定引数の取得
    Dim strKind       '入力種類(1=届時刻,2=完了時刻)
    Dim iLine         '入力行
    Dim strRequest    '戻り先
    strKind=Session.Contents("editkind")
    iLine=CInt(Session.Contents("editline"))
    strRequest=Session.Contents("request")

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
    If strKind="1" Then
        sTranID = "IM10"
    Else
        sTranID = "IM11"
    End If
    ' 処理区分
    Const sSyori = "R"
    ' 送信場所
    Const sPlace = ""

    ' セッションのチェック
    CheckLogin "ms-impentry.asp&kind=2"
    sSosin = Trim(Session.Contents("userid"))

	' タイトル取得
	strTitle = Trim(Request.form("title"))

    ' エラーフラグのクリア
    bError = false

    If Not bError Then
        'トランザクションファイル作成
        anyTmp=Split(strData(iLine-1),",")
        ' テンポラリファイル名を作成して、セッション変数に設定
        Dim sIMxx, iSeqNo, strFileName_01, sTran, sTusin, sDate
        'シーケンス番号
        iSeqNo = GetDailyTransNo
        '通信日時取得
        sTusin  = SetTusinDate
        sDate = Trim(Request.form("Year")) 
        sDate = sDate & Right("0" & Trim(Request.form("Month")),2)
        sDate = sDate & Right("0" & Trim(Request.form("Day")),2)
        sDate = sDate & Right("0" & Trim(Request.form("Hour")),2)
        sDate = sDate & Right("0" & Trim(Request.form("Min")),2)

        If strKind="1" Then
            sIMxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
                    sSosin & "," & sPlace & "," & anyTmp(4) & "," &  anyTmp(3) & "," & _
                    anyTmp(1) & "," & anyTmp(0) & "," & sDate & "," & sSosin
        Else
            sIMxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
                    sSosin & "," & sPlace & "," & anyTmp(4) & "," &  anyTmp(3) & "," & _
                    anyTmp(1) & "," & anyTmp(0) & "," & sDate & "," & sSosin
        End If
        sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
        strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
        ti.WriteLine sIMxx
        ti.Close

        sLogDate = Trim(Request.form("Year")) & "/"
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Month")),2) & "/"
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Day")),2) & " "
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Hour")),2) & ":"
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Min")),2)
        strOption = anyTmp(0) & "/" & sLogDate & "," & "入力内容の正誤:0(正しい)"

        ' テンポラリファイル更新
        strTemp=Left(sDate,4) & "/" & Mid(sDate,5,2) & "/" & Mid(sDate,7,2) & " " & Mid(sDate,9,2) & ":" & Mid(sDate,11,2)
        If strKind="1" Then
            anyTmp(44)=strTemp
        Else
            anyTmp(45)=strTemp
        End If
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
        For i=1 To LineNo
            If i<>CInt(iLine) Then
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

        ' 海貨入力項目選択
        If strKind="1" Then
            WriteLog fs, "","輸出入業務支援-輸入実入りコンテナ倉庫到着時刻入力","10", strOption
        Else
            WriteLog fs, "2107","輸入コンテナ照会-デバンニング完了時刻入力","10", strOption
        End If

        If strRequest="ms-impdetail.asp" Then
            strTemp=strRequest & "?line=" & iLine
        Else
            strTemp=strRequest
        End If

        ' 戻り画面へリダイレクト
        Response.Redirect strTemp
        Response.End
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
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/imprikuun.gif" width="506" height="73"></td>
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
				<%=strRoute%> &gt; 時刻入力
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
    strOption = anyTmp(1) & "/" & sLogDate & "," & "入力内容の正誤:1(誤り)"
    ' 海貨入力項目選択
    If strKind="1" Then
        WriteLog fs, "","輸出入業務支援-輸入実入りコンテナ倉庫到着時刻入力","10", strOption
    Else
        WriteLog fs, "2107","輸入コンテナ照会-デバンニング完了時刻入力","10", strOption
    End If
%>
      <br><br>
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
    If strRequest="ms-impdetail.asp" Then
        strTemp=strRequest & "?line=" & iLine
    Else
        strTemp=strRequest
    End If
    DispMenuBarBack strTemp
%>
</body>
</html>
<%

%>