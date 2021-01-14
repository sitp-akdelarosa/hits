<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
''    ' セッションのチェック
''    CheckLogin "expentry.asp"
    ' エラーフラグのクリア
    bError = false

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '入力画面を記憶
    Session.Contents("findcsv")="no"    ' 直接入力であることを記憶

    ' 指定引数の取得
    Dim strCntnrNo,strCntnrNoLog
    Dim strBookingNo,strBookingNoLog
' 2009/05/09 add-s 港条件追加
    Dim strUserCodeEx,strUserCodeExNoLog
' 2009/05/09 add-s 港条件追加
    strCntnrNo = UCase(Trim(Request.QueryString("cntnrno")))
    strBookingNo = UCase(Trim(Request.QueryString("booking")))
' 2009/05/09 add-s 港条件追加
    strUserCodeEx = UCase(Trim(Request.QueryString("portcode")))
	strUserCodeExNoLog = strUserCodeEx
' 2009/05/09 add-s 港条件追加
	strCntnrNoLog = strCntnrNo
	strBookingNoLog = strBookingNo
    If strCntnrNo="" And strBookingNo="" Then
        ' 引数指定のないとき エラーメッセージを表示
        bError = true
        strError = "参照したいコンテナNo.又は、Booking Noのうち、<br>一項目は入力してください。"
        strOption = "," & "," & ",入力内容の正誤:1(誤り)"
        ' 引数指定のないとき サンプル画面を表示
        Response.Redirect "explist.html"
        Responce.End
    Else
        ' テンポラリファイル名を作成して、セッション変数に設定
        Dim strFileName
        strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
        Session.Contents("tempfile")=strFileName

' 2009/05/09 add-s 港条件追加
        Session.Contents("usercodeex")=strUserCodeEx     ' 参照ユーザーコードを記憶
' 2009/05/09 add-e 港条件追加

        ' コンテナ情報レコードの取得
        ConnectSvr conn, rsd
        sWhere = ""
'☆☆☆ Add_S  by nics 200902改造
        dim bWriteFile
        bWriteFile = 0
        ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
        strFileName="./temp/" & strFileName
        ' テンポラリファイルのOpen
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
'☆☆☆ Add_E  by nics 200902改造
        If strBookingNo<>"" Then        ' Booking番号の入力が優先
            strInput = "," & "入力内容," & strBookingNo
            strOption = "入力方法分類,3(Booking番号1つ)" & strInput

            Session.Contents("findkey")=strBookingNo     ' 参照Keyを記憶
'☆☆☆ Mod_S  by nics 200902改造
'            iCanma = InStr(strBookingNo,",")
'            Do While iCanma>0
'                strOption = "入力方法分類,4(Booking番号複数)" & strInput
'                sTemp = Trim(Left(strBookingNo,iCanma-1))
'                strBookingNo = Right(strBookingNo,Len(strBookingNo)-iCanma)
'                If sWhere<>"" Then
'                    sWhere = sWhere & " Or ExportCont.BookNo='" & sTemp & "'"
'                Else
'                    sWhere = "ExportCont.BookNo='" & sTemp & "'"
'                End If
'                iCanma = InStr(strBookingNo,",")
'            Loop
'            If sWhere<>"" Then
'                sWhere = sWhere & " Or ExportCont.BookNo='" & Trim(strBookingNo) & "'"
'            Else
'                sWhere = "ExportCont.BookNo='" & Trim(strBookingNo) & "'"
'            End If
'☆☆☆
            Do While strBookingNo <> ""
                iCanma = InStr(strBookingNo, ",")
                If iCanma > 0 Then
                    strOption = "入力方法分類,4(Booking番号複数)" & strInput
                    sTemp = Left(strBookingNo, iCanma-1)
                    strBookingNo = Mid(strBookingNo, iCanma+1)
                Else
                    sTemp = strBookingNo
                    strBookingNo = ""
                End If
' 2009/05/09 mod-s 港条件追加/SQLインジェクション対応
'                sWhere = "ExportCont.BookNo='" & Trim(sTemp) & "'"
'                bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                bRtn = ChkSQLInjectionBookNo(sTemp)
                If bRtn Then
                    sWhere = "ExportCont.BookNo='" & Trim(sTemp) & "' and Container.UserCode='" & Trim(strUserCodeEx) & "'"
                    bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                End If
' 2009/05/09 mod-e 港条件追加/SQLインジェクション対応
            Loop
'☆☆☆ Mod_E  by nics 200902改造
            Session.Contents("findkind")="Booking"       ' 参照モード
        Else
            strInput = "," & "入力内容," & strCntnrNo
            strOption = "入力方法分類,0(コンテナNo.1つ)" & strInput

            Session.Contents("findkey")=strCntnrNo       ' 参照Keyを記憶
'☆☆☆ Mod_S  by nics 200902改造
'            iCanma = InStr(strCntnrNo,",")
'            Do While iCanma>0
'                strOption = "入力方法分類,1(コンテナNo.複数)" & strInput
'                sTemp = Trim(Left(strCntnrNo,iCanma-1))
'                strCntnrNo = Right(strCntnrNo,Len(strCntnrNo)-iCanma)
'                If sWhere<>"" Then
'                    sWhere = sWhere & " Or ExportCont.ContNo='" & sTemp & "'"
'                Else
'                    sWhere = "ExportCont.ContNo='" & sTemp & "'"
'                End If
'                iCanma = InStr(strCntnrNo,",")
'            Loop
'            If sWhere<>"" Then
'                sWhere = sWhere & " Or ExportCont.ContNo='" & Trim(strCntnrNo) & "'"
'            Else
'                sWhere = "ExportCont.ContNo='" & Trim(strCntnrNo) & "'"
'            End If
'☆☆☆
            Do While strCntnrNo <> ""
                iCanma = InStr(strCntnrNo, ",")
                If iCanma > 0 Then
                    strOption = "入力方法分類,1(コンテナNo.複数)" & strInput
                    sTemp = Left(strCntnrNo, iCanma-1)
                    strCntnrNo = Mid(strCntnrNo, iCanma+1)
                Else
                    sTemp = strCntnrNo
                    strCntnrNo = ""
                End If
' 2009/05/09 mod-s 港条件追加/SQLインジェクション対応
'                sWhere = "ExportCont.ContNo='" & Trim(sTemp) & "'"
'                bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                bRtn = ChkSQLInjectionCntnrNo(sTemp)
                If bRtn Then
                    sWhere = "ExportCont.ContNo='" & Trim(sTemp) & "' and Container.UserCode='" & Trim(strUserCodeEx) & "'"
                    bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                End If
' 2009/05/09 mod-e 港条件追加/SQLインジェクション対応
            Loop
'☆☆☆ Mod_E  by nics 200902改造
            Session.Contents("findkind")="Cntnr"         ' 参照モード
        End If

'☆☆☆ Del_S  by nics 200902改造
'        ' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
'        strFileName="./temp/" & strFileName
'        ' テンポラリファイルのOpen
'        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
'
'        bWriteFile = SerchExpCntnr(conn, rsd, ti, sWhere)
'☆☆☆ Del_E  by nics 200902改造

        ti.Close
        conn.Close

        ' Tempファイル属性設定
        SetTempFile "EXPORT"

        If bWriteFile = 0 Then
            ' 該当レコードないとき
            bError = true
            strError = "指定条件に該当するコンテナはありませんでした。"
            strOption = "入力内容の正誤:1(誤り)"
        Else
            strOption = "入力内容の正誤:0(正しい)"
        End If

    End If

	Dim iWrkNum
	If strBookingNoLog="" Then
		iWrkNum = 11
		Do While InStr(strCntnrNoLog,",")>0
			strCntnrNoLog = Left(strCntnrNoLog,InStr(strCntnrNoLog,",")-1) & _
							"/" & Right(strCntnrNoLog,Len(strCntnrNoLog)-InStr(strCntnrNoLog,",")) & _
							"/" & strUserCodeExNoLog
		Loop
		strOption = strCntnrNoLog & "," & strOption
	Else
		iWrkNum = 12
		Do While InStr(strBookingNoLog,",")>0
			strBookingNoLog = Left(strBookingNoLog,InStr(strBookingNoLog,",")-1) & _
							"/" & Right(strBookingNoLog,Len(strBookingNoLog)-InStr(strBookingNoLog,",")) & _
							"/" & strUserCodeExNoLog
		Loop
		strOption = strBookingNoLog & "," & strOption
	End If

    ' 輸出コンテナ照会
    WriteLog fs, "1001","輸出コンテナ照会(外部)",iWrkNum, strOption

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
<!--
function FancBack()
{
        window.history.back();
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>仕出地情報</b></td>
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
    DispMenuBarBack "JavaScript:FancBack()"
%>
</body>
</html>

<%
    Else
        If bWriteFile = 1 Then
            '戻り画面種別を記憶
            Session.Contents("dispreturn")=0

            ' 詳細画面へリダイレクト
            Response.Redirect "expdetail.asp?line=1"     '輸出コンテナ詳細
        Else
            '戻り画面種別を記憶
            Session.Contents("dispreturn")=0
            ' 一覧画面へリダイレクト
            Response.Redirect "explist.asp"             '輸出コンテナ一覧
        End If
    End If
%>
