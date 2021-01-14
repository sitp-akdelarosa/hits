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

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	Dim bChkboxFlag
	bChkboxFlag = false

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

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim sEXxx, iSeqNo, strFileName_01, sTran, sTusin, sDate
    'シーケンス番号
    iSeqNo = GetDailyTransNo
    '通信日時取得
    sTusin  = SetTusinDate


	'引当可否設定
	If Request.Form("ok")<>"" Then

      sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
      strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
      Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

	  Set titmp=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

	  For i=1 To CInt(Request.Form("allline"))

        'トランザクションファイル作成
		If Request.Form("check" & i)<>"" Then
	        anyTmp=Split(strData(i-1),",")

			If anyTmp(24)<>"" Then
				sDate = Left(anyTmp(24),4) & Mid(anyTmp(24),6,2) & Mid(anyTmp(24),9,2)
			End If

	        sEXxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
	                sSosin & "," & sPlace & "," & anyTmp(25) & "," & anyTmp(4) & "," & anyTmp(3) & "," &_
					anyTmp(0) & "," & anyTmp(23) & "," & anyTmp(14) & "," & "1" & "," &_
					anyTmp(20) & "," & sDate

	        ti.WriteLine sEXxx

			bChkboxFlag = true
		End If

        ' テンポラリファイル更新
		If Request.Form("check" & i)="" Then
            titmp.WriteLine strData(i-1)
        Else
            strTemp=anyTmp(0)
			anyTmp(26) = "1"
            For j=1 To UBound(anyTmp)
                strTemp=strTemp & "," & anyTmp(j)
            Next
            titmp.WriteLine strTemp
        End If

	  Next

      titmp.Close
      ti.Close

	  If bChkboxFlag Then
		  WriteLog fs, "a108","空コンピックアップシステム-港運用情報一覧","10", ","
		  Response.Redirect "picklist.asp?kind=4"
	  End If

	Else

	  Dim sLineAry
	  For i=1 To CInt(Request.Form("allline"))
		If Request.Form("check" & i)="on" Then
			If sLineAry="" Then
				sLineAry = i
			Else
				sLineAry = sLineAry & "," & i
			End If
			bChkboxFlag = true
		End If
	  Next

	  If bChkboxFlag Then
		  Session.Contents("lines") = sLineAry
		  Response.Redirect "picklist-input.asp"
	  End If

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
          <td nowrap><b>
            空コンピックアップ情報一覧（港運用）</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
	 	<BR>

<%
	strError = "確認／変更チェックボックスのチェックが付いていません。"
    ' エラーメッセージの表示
    DispErrorMessage strError 
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
    strTemp = "picklist.asp?kind=4"
    DispMenuBarBack strTemp
%>
</body>
</html>
<%

%>
