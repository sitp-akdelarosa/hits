<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->

<%
	' 寄港地のＭＡＸ値
	Const KIKOUTI = 30

    ' セッションのチェック
    CheckLogin "nyuryoku-in1.asp"

    ' エラーフラグのクリア
    bError = false

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

	' 入力港名のチェック
	'新規入力時のポートコード
    Dim strPort
	strPort =  UCase(Trim(Request.Form("port")))

	ConnectSvr conn, rsd
	sql = "SELECT FullName FROM mPort WHERE PortCode='" & strPort & "'"
	'SQLを発行して港コードマスターを検索
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		strPortName = Trim(rsd("FullName"))
		strOption = "入力内容の正誤:0(正しい)"
	Else
		' 該当レコードのないとき エラーメッセージを表示
		bError = true
		strError = "該当する港コードが有りません。"
		strOption = "入力内容の正誤:1(誤り)"
	End If
	rsd.Close

	Dim sAdate, sTdate, sDdate, sCdate, sRdate
	'各時刻設定
		'着岸予定時刻
		sAdate = ""
		If Request.Form("ayear") <> "" Then
		    sAdate = SetDateTime(Request.Form("ayear"), Request.Form("amonth"), Request.Form("aday"), _ 
		                         GetNumStr(Request.Form("ahour"), 2), GetNumStr(Request.Form("amin"), 2))
		End If
		'着岸完了時刻
		sTdate = ""
		If Request.Form("tyear") <> "" Then
		    sTdate = SetDateTime(Request.Form("tyear"), Request.Form("tmonth"), Request.Form("tday"), _ 
		                         GetNumStr(Request.Form("thour"), 2), GetNumStr(Request.Form("tmin"), 2))
		End If
		'離岸完了時刻
		sDdate = ""
		If Request.Form("dyear") <> "" Then
		    sDdate = SetDateTime(Request.Form("dyear"), Request.Form("dmonth"), Request.Form("dday"), _ 
		                         GetNumStr(Request.Form("dhour"), 2), GetNumStr(Request.Form("dmin"), 2))
		End If
		'着岸Long Schedule
		sCdate = ""
		If Request.Form("cyear") <> "" Then
		    sCdate = SetDateTime(Request.Form("cyear"), Request.Form("cmonth"), Request.Form("cday"), _ 
                                 "23", "59")
'		                         GetNumStr(Request.Form("chour"), 2), GetNumStr(Request.Form("cmin"), 2))
		End If
		'離岸Long Schedule
		sRdate = ""
		If Request.Form("ryear") <> "" Then
		    sRdate = SetDateTime(Request.Form("ryear"), Request.Form("rmonth"), Request.Form("rday"), _ 
                                 "23", "59")
'		                         GetNumStr(Request.Form("rhour"), 2), GetNumStr(Request.Form("rmin"), 2))
		End If

	If not bError Then
	    ' 指定引数の取得(指定行)
	    Dim iLine, sIn1, sIn2, sInpFlg
		Dim sText(35) 

	    ' 詳細表示行のデータの取得

	    Dim iKensu		'表示件数(画面表示件数)
	    Dim LineNo		'ファイルのラインカウンタ
		Dim iDelLine	'削除する行番号

	    LineNo=0
	    Do While Not ti.AtEndOfStream
	        anyTmp=Split(ti.ReadLine,",")
	        LineNo=LineNo+1
			Select Case LineNo
				Case 1
					iKensu = anyTmp(7) + 1
					If iKensu > KIKOUTI Then
						bError = true
						strError = "寄港地情報の入力がＭＡＸ値を越えました。"
						strOption = "入力内容の正誤:1(誤り)"
						Exit Do
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
					sText(LineNo) = anyTmp(0) &  "," & anyTmp(1) &  "," & _
								    anyTmp(2) &  "," & anyTmp(3) &  "," & _
								    anyTmp(4) &  "," & anyTmp(5) &  "," & _
							    	anyTmp(6) &  "," & anyTmp(7) 
			End Select
	    Loop

		'港名称重複チェック
		For i = 3 to LineNo - 1
			anyTmp=Split(sText(i),",")
			If Trim(anyTmp(0)) = strPort Then
				bError = true
				strError = "寄港地情報が既に登録されています。"
				strOption = "入力内容の正誤:1(誤り)"
			End If
		Next 

		If not bError Then
		    ti.Close
			LineNo = LineNo + 1
			sText(LineNo) = strPort &  "," & strPortName &  "," & _
						    saDate    &  "," & sTdate    &  ",," & _
					    	sDdate    &  "," & _
					    	sCdate &  "," & sRdate

		'順番並び替えの処理を行う(小西さんの要望で、コメント化 2002/02/27)
		'*** Start M.Hayashi ****
'			Dim sBefDate
'			Dim sAftDate
'		    Dim sWkText
'		    Dim bSwap
'		    For i = 3 to LineNo - 1
'				anyTmp=Split(sText(i),",")
'				sBefDate = anyTmp(2)
'				For j = (i + 1) To LineNo
'		            anyTmp=Split(sText(j),",")
'				    sAftDate = anyTmp(2)
'		            bSwap = FALSE
'		            If sAftDate <> "" Then
'					  If sBefDate = "" Then
'		                bSwap = TRUE
'		              Else
'		                If sBefDate > sAftDate Then
'		                  bSwap = TRUE
'						End If
'		              End IF
'		            End If
'		            If bSwap = TRUE Then
'		              sWkText = sText(i)
'		              sText(i) = sText(j)
'		              sText(j) = sWkText
'		            End IF
'				Next 
'			Next
		'*** End   M.Hayashi ****

		    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
		    For i = 1 to LineNo
				ti.WriteLine sText(i)
		    Next
		    ti.Close
		End If
	End If

	sCdate = Left(sCdate,10)
	sRdate = Left(sRdate,10)
    ' 船社入力新規入力
	WriteLog fs, "3004","船社／ターミナル入力-本船動静入力","11", strPort & "/" & sAdate & "/" & sTdate & "/" & sDdate & "/" & sCdate & "/" & sRdate & "," & strOption

%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから一覧画面--------------------------->
<%
	If not bError Then
%>
<!-------------登録画面更新処理終わり--------------------------->
</body>
<SCRIPT LANGUAGE="JavaScript">
	window.location.replace("nyuryoku-port.asp");
</SCRIPT>
<%	Else	%>
<SCRIPT LANGUAGE="JavaScript">
function FancBack()
{
        window.history.back();
}
</SCRIPT>
<!-------------ここから照会エラー画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/nyuryoku-s.gif" width="506" height="73"></td>
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
                    <td nowrap><b>本船動静入力　</b></td>
              <td><img src="gif/hr.gif"></td>
            </tr>
          </table>
<br>     
      <table>
        <tr>
          <td>
<%
    ' メッセージの表示
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
'    DispMenuBarBack "nyuryoku-new.asp"
%>
<map name="map"> 
  <area shape="poly" coords="20,0,152,0,134,22,0,22" href="JavaScript:FancBack()">
  <area shape="poly" coords="154,0,136,22,284,22,284,0" href="http://www.hits-h.com/index.asp">
</map>

<body>
<%	End If %>
</html>
