<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-in1.asp"

    ' エラーフラグのクリア
    bError = false

    ' 入力フラグのクリア
    bInput = true

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' 引数指定のないとき
        strFileName="test.csv"
    End If
    strFileName="./temp/" & strFileName
    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 指定引数の取得(指定行)
    Dim iLine, sIn1, sIn2, sInpFlg
	Dim sText(35) 

    iLine = Trim(Request.QueryString("line"))

    ' 詳細表示行のデータの取得

    Dim iKensu		'表示件数(削除後)
    Dim LineNo		'ファイルのラインカウンタ
    Dim iHitNo		'一致するファイル行数
	Dim iDelLine	'削除する行番号
	Dim sPortName	'ログ用港名

    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
		Select Case LineNo
			Case 1
				iKensu = anyTmp(7) - 1
				If Cint(iKensu) <> 0 Then
		    		iHitNo = 2 +  Cint(iLine)
				End If
				sText(LineNo) = anyTmp(0) &  "," & _
							    anyTmp(1) &  "," & _
							    anyTmp(2) &  "," & _
							    anyTmp(3) &  "," & _
							    anyTmp(4) &  "," & _
							    anyTmp(5) &  "," & _
							    anyTmp(6) &  "," & iKensu
			Case 2
				sText(LineNo) = iKensu
			Case Else
				If iKensu = 0 Then
					Exit Do
				End If

		        If LineNo = iHitNo Then
					iDelLine = LineNo
					sPortName = anyTmp(1)
				Else
					sText(LineNo) = anyTmp(0) &  "," & _
								    anyTmp(1) &  "," & _
								    anyTmp(2) &  "," & _
								    anyTmp(3) &  "," & _
								    anyTmp(4) &  "," & _
							    	anyTmp(5) &  "," & _
							    	anyTmp(6) &  "," & _
							    	anyTmp(7)
		        End If
		End Select
    Loop
    ti.Close


	sBk = Server.MapPath(strFileName)
	sTemp  = strFileName & ".tmp" 	'Server.MapPath(strFileName)
	fs.deletefile sBk, True				'一度削除
	ti  = Server.MapPath(sTemp)
    Set ti=fs.OpenTextFile(Server.MapPath(sTemp),2,True)
    For i = 1 to LineNo

		If iKensu <> 0 Then
			If iDelLine <> i Then
				ti.WriteLine sText(i)
			End If
		Else
			ti.WriteLine sText(i)
			If i = 2 Then
				Exit For
			End If
		End If
    Next
	ti  = Server.MapPath(sTemp)
	sBk = Server.MapPath(strFileName)
	fs.MoveFile ti, sBk



    ' 本船動静削除
    WriteLog fs, "3004","船社／ターミナル入力-本船動静入力", "13", sPortName & ","

%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body>
<!-------------ここから一覧画面--------------------------->
<!-------------登録画面更新処理終わり--------------------------->
</body>
<SCRIPT LANGUAGE="JavaScript">
	window.location.replace("nyuryoku-port.asp");
</SCRIPT>
</html>
