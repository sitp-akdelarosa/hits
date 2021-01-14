<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ImpCom.inc"-->

<%
''    ' セッションのチェック
''    CheckLogin "impentry.asp"

    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '入力画面を記憶
    Session.Contents("findcsv")="no"    ' 直接入力であることを記憶

    ' 指定引数の取得
    Dim strCntnrNo,strCntnrNoLog
    Dim strBLNo,strBLNoLog
    strCntnrNo = UCase(Trim(Request.QueryString("cntnrno")))
    strBLNo = UCase(Trim(Request.QueryString("blno")))
	strCntnrNoLog = strCntnrNo
	strBLNoLog = strBLNo
    If strCntnrNo="" And strBLNo="" Then
        ' 引数指定のないとき エラーメッセージを表示
        bError = true
        strError = "参照したいコンテナNo.又は、B/L Noのうち、<br>一項目は入力してください。"
        strOption = "," & "," & "入力内容の正誤:1(誤り)"
        ' 引数指定のないとき サンプル画面を表示
        Response.Redirect "implist.html"
        Responce.End
    Else
        ' テンポラリファイル名を作成して、セッション変数に設定
        Dim strFileName
        strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
        Session.Contents("tempfile")=strFileName

        ' コンテナ情報レコードの取得
        ConnectSvr conn, rsd
        sWhere = ""
        If strBLNo<>"" Then        ' BL番号の入力が優先
            strInput = "," & "入力内容," & strBookingNo
            strOption = "入力方法分類,3(BL番号1つ)" & strInput

            Session.Contents("findkey")=strBLNo         ' 参照Keyを記憶
            iCanma = InStr(strBLNo,",")
            Do While iCanma>0
                strOption = "入力方法分類,4(BL番号複数)" & strInput
                sTemp = Trim(Left(strBLNo,iCanma-1))
                strBLNo = Right(strBLNo,Len(strBLNo)-iCanma)
                If sWhere<>"" Then
                    sWhere = sWhere & " Or ImportCont.BLNo='" & sTemp & "'"
                Else
                    sWhere = "ImportCont.BLNo='" & sTemp & "'"
                End If
                iCanma = InStr(strBLNo,",")
            Loop
            If sWhere<>"" Then
                sWhere = sWhere & " Or ImportCont.BLNo='" & Trim(strBLNo) & "'"
            Else
                sWhere = "ImportCont.BLNo='" & Trim(strBLNo) & "'"
            End If
            Session.Contents("findkind")="Blno"       ' 参照モード
        Else
            strInput = "," & "入力内容," & strCntnrNo
            strOption = "入力方法分類,0(コンテナNo.1つ)" & strInput

            Session.Contents("findkey")=strCntnrNo       ' 参照Keyを記憶
            iCanma = InStr(strCntnrNo,",")
            Do While iCanma>0
                strOption = "入力方法分類,1(コンテナNo.複数)" & strInput
                sTemp = Trim(Left(strCntnrNo,iCanma-1))
                strCntnrNo = Right(strCntnrNo,Len(strCntnrNo)-iCanma)
                If sWhere<>"" Then
                    sWhere = sWhere & " Or ImportCont.ContNo='" & sTemp & "'"
                Else
                    sWhere = "ImportCont.ContNo='" & sTemp & "'"
                End If
                iCanma = InStr(strCntnrNo,",")
            Loop
            If sWhere<>"" Then
                sWhere = sWhere & " Or ImportCont.ContNo='" & Trim(strCntnrNo) & "'"
            Else
                sWhere = "ImportCont.ContNo='" & Trim(strCntnrNo) & "'"
            End If
            Session.Contents("findkind")="Cntnr"         ' 参照モード
        End If

        ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
        strFileName="./temp/" & strFileName
        ' テンポラリファイルのOpen
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

        bWriteFile = SerchImpCntnr(conn, rsd, ti, sWhere)

        ti.Close
        conn.Close

        ' Tempファイル属性設定
        SetTempFile "IMPORT"

        If bWriteFile = 0 Then
            ' 該当レコードないとき
            bError = true
            strError = "No container that corresponded to a specified condition."
            strOption = "入力内容の正誤:1(誤り)"
        Else
            strOption = "入力内容の正誤:0(正しい)"
        End If

    End If

	Dim iWrkNum
	If strBLNoLog="" Then
		iWrkNum = 11
		Do While InStr(strCntnrNoLog,",")>0
			strCntnrNoLog = Left(strCntnrNoLog,InStr(strCntnrNoLog,",")-1) & _
							"/" & Right(strCntnrNoLog,Len(strCntnrNoLog)-InStr(strCntnrNoLog,","))
		Loop
		strOption = strCntnrNoLog & "," & strOption
	Else
		iWrkNum = 12
		Do While InStr(strBLNoLog,",")>0
			strBLNoLog = Left(strBLNoLog,InStr(strBLNoLog,",")-1) & _
							"/" & Right(strBLNoLog,Len(strBLNoLog)-InStr(strBLNoLog,","))
		Loop
		strOption = strBLNoLog & "," & strOption
	End If

    ' 輸入コンテナ照会
    WriteLog fs, "2301","輸入コンテナ照会",iWrkNum, strOption

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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから照会エラー画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="../gif/shokait.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
'	DisplayCodeListButton
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
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>Container Information (Imp.)</b></td>
          <td><img src="../gif/hr.gif"></td>
        </tr>
      </table>
		<BR>
      <table>
        <tr>
          <td>
<%
    ' エラーメッセージの表示
    DispErrorMessage strError
%>
          </td></tr>
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
<!-------------照会エラー画面終わり--------------------------->
<%
    DispMenuBarBack "impentry.asp"
%>
</body>
</html>

<%
    Else
        If bWriteFile = 1 Then
            '戻り画面種別を記憶
            Session.Contents("dispreturn")=0
            ' 詳細画面へリダイレクト
            Response.Redirect "impdetail.asp?line=1"    '輸入コンテナ詳細
        Else
            '戻り画面種別を記憶
            Session.Contents("dispreturn")=0
            ' 一覧画面へリダイレクト
            Response.Redirect "implist.asp"             '輸入コンテナ一覧
        End If
    End If
%>
