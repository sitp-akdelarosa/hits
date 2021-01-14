<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-in1.asp"

    ' エラーフラグのクリア
    bError = false

    ' 入力フラグのクリア
    bInput = true

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' テンポラリファイル名を作成して、セッション変数に設定
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' 指定引数の取得
    Dim strCallSign
    Dim strVoyage
    strCallSign = UCase(Trim(Request.QueryString("callsign")))
    strVoyage = UCase(Trim(Request.QueryString("voyage")))

	If InStr(strVoyage,",")<>0 Then
        ' カンマ入力時エラー
        bError = true
        strError = "Voyage No.に半角カンマは使用しないで下さい。"
        strOption = strCallSign & "/" & strVoyage & "," & "入力内容の正誤:1(誤り)"
    End If

    If strCallSign="" Or strVoyage="" Then
        If strCallSign<>"" Or strVoyage<>"" Then
            ' 入力が片方だけのとき エラーメッセージを表示
            bError = true
            strError = "入力が間違っています。"
            strOption = strCallSign & "/" & strVoyage & "," & "入力内容の正誤:1(誤り)"
        Else
            bInput = false
        End If
    End If

    If bInput And Not bError Then
        ' 入力コールサインのチェック
        ConnectSvr conn, rsd
        sql = "SELECT FullName, ShipLine FROM mVessel WHERE VslCode='" & strCallSign & "'"
        'SQLを発行して船名マスターを検索
        rsd.Open sql, conn, 0, 1, 1
        If Not rsd.EOF Then
            strVesselName = Trim(rsd("FullName"))
            strShipLine = Trim(rsd("ShipLine"))
            strOption = strCallSign & "/" & strVoyage & "," & "入力内容の正誤:0(正しい)"
        Else
            ' 該当レコードのないとき エラーメッセージを表示
            bError = true
            strError = "コールサインが間違っています。"
            strOption = strCallSign & "/" & strVoyage & "," & "入力内容の正誤:1(誤り)"
        End If
        rsd.Close
        If Not bError Then
            ' 船社名の取得
            sql = "SELECT FullName FROM mShipLine WHERE ShipLine='" & strShipLine & "'"
            'SQLを発行して船名マスターを検索
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                strShipLineName = Trim(rsd("FullName"))
            End If
            rsd.Close

            Dim strPortData()

            ' SQLを発行して本船動静を検索
            sql = "SELECT VoyCtrl, DsVoyage, LdVoyage FROM VslSchedule " & _
                  "WHERE VslCode='" & strCallSign & "' And " & _
                  "(DsVoyage='" & strVoyage & "' Or LdVoyage='" & strVoyage & "')"
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                iVoyCtrl = rsd("VoyCtrl")
                strDsVoyage = Trim(rsd("DsVoyage"))
                strLdVoyage = Trim(rsd("LdVoyage"))
                rsd.Close
                ' 本船動静情報レコードの作製(着岸予定時刻が入っているものを先に読む)
                strVslSchdule = strShipLine & "," & strShipLineName & "," & strCallSign & "," & strVesselName & "," & _
								iVoyCtrl & "," & strDsVoyage & "," & strLdVoyage
                ' SQLを発行して本船寄港地を検索(小西さんの要望で、寄港地順に 2002/02/27)
'               sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong, mPort.FullName " & _
'                     "FROM VslPort, mPort WHERE VslPort.VslCode='" & strCallSign & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
'                     " And mPort.PortCode=*VslPort.PortCode And VslPort.ETA is NOT Null ORDER BY VslPort.ETA "
                sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong, mPort.FullName " & _
                      "FROM VslPort, mPort WHERE VslPort.VslCode='" & strCallSign & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
                      " And mPort.PortCode=*VslPort.PortCode ORDER BY VslPort.CallSeq "
                rsd.Open sql, conn, 0, 1, 1
                iRecCount=0
				iSeq = 1
                Do While Not rsd.EOF
                    ' 寄港地情報レコードの作製
                    strRec = Trim(rsd("PortCode")) & "," & Trim(rsd("FullName")) & "," & _
							 DispDateTime(rsd("ETA"),0) & "," & DispDateTime(rsd("TA"),0) & ","  & _
							 DispDateTime(rsd("ETD"),0) & "," & DispDateTime(rsd("TD"),0) & ","  & _
							 DispDateTime(rsd("ETALong"),0) & "," & DispDateTime(rsd("ETDLong"),0)
                    ReDim Preserve strPortData(iRecCount)
                    strPortData(iRecCount) = strRec
                    iRecCount=iRecCount + 1
					iSeq = iSeq + 1
                    rsd.MoveNext
                Loop
' 				rsd.Close
'01/12/22 ADD (小西さんの要望で、寄港地順にしたため不要に 2002/02/27)
'                ' 本船動静情報レコードの作製(着岸予定時刻が入っていないものを読む)
'                ' SQLを発行して本船寄港地を検索
'                sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong, mPort.FullName " & _
'                      "FROM VslPort, mPort WHERE VslPort.VslCode='" & strCallSign & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
'                      " And mPort.PortCode=*VslPort.PortCode And VslPort.ETA is Null"
'                rsd.Open sql, conn, 0, 1, 1
'                Do While Not rsd.EOF
'                    ' 寄港地情報レコードの作製
'                    strRec = Trim(rsd("PortCode")) & "," & Trim(rsd("FullName")) & "," & _
'							 DispDateTime(rsd("ETA"),0) & "," & DispDateTime(rsd("TA"),0) & ","  & _
'							 DispDateTime(rsd("ETD"),0) & "," & DispDateTime(rsd("TD"),0) & ","  & _
'							 DispDateTime(rsd("ETALong"),0) & "," & DispDateTime(rsd("ETDLong"),0)
'                    ReDim Preserve strPortData(iRecCount)
'                    strPortData(iRecCount) = strRec
'                    iRecCount=iRecCount + 1
'					iSeq = iSeq + 1
'                    rsd.MoveNext
'                Loop

            Else
                ' 本船動静情報レコードの作製
                strVslSchdule = strShipLine & "," & strShipLineName & "," & strCallSign & "," & strVesselName & ",,"  & _
								strVoyage & "," & strVoyage
                iRecCount=0
            End If
            rsd.Close
            ' 検索データを一時ファイルに出力
            strFileName="./temp/" & strFileName
            ' テンポラリファイルのOpen
            Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

            ti.WriteLine strVslSchdule & "," & iRecCount

            ti.WriteLine iRecCount
            For iCount=0 To iRecCount - 1
                ti.WriteLine strPortData(iCount)
            Next

            ti.Close
        End If
        conn.Close
    End If

    If bError Or Not bInput Then
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
<!-------------ここから船名入力画面--------------------------->
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

<table border=0><tr><td align=left>
  <table>
                  <tr>
                    
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                
                  <td nowrap><b>本船動静入力</b></td>
                   <td><img src="gif/hr.gif"></td>
 </tr>
</table>
 <center>             
	  <table>
	   <tr>
	                <td nowrap>対象となる本船に関する下記の情報を入力の上、<BR>送信ボタンをクリックして下さい。</td>
          </tr>
		</table>
            	<FORM NAME="con" action="nyuryoku-in1.asp">
                  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                    <tr> 
                      <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF"> 
                        コールサイン</font></b></td>
                      <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=120>
								<input type=text name=callsign value="<%=strCallSign%>" size=10 maxlength=7>
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
								<font size=1 color="#2288ff">[ 半角英数 ]</font>
							</td>
						  </tr>
						</table>
				                    	
                      </td>
                    </tr>
                    <tr> 
                      <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">Voyage 
                        No.</font></b></td>
                      <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=120>
								<input type=text name=voyage value="<%=strVoyage%>" size=12 maxlength=12>
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
								<font size=1 color="#2288ff">[ 半角英数 ]</font>
							</td>
						  </tr>
						</table>
                      </td>
                    </tr>
                  </table>
                  <br>
			            <INPUT TYPE=submit VALUE=" 送  信 " name="送信">
			<BR>

<%
        ' エラーメッセージの表示
        If bError Then
            DispErrorMessage strError
       End If
%>
			<BR>
</center>
                  <table>
                    <tr> 
                      <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                      <td nowrap><b>CSVファイル転送</b></td>
                      <td><img src="gif/hr.gif"></td>
                    </tr>
                  </table>
<center>
<table border="0" cellspacing="1" cellpadding="2">


          <tr> 
              <td> 
                <p>情報をファイル転送する場合はここをクリック</p>
              </td>
              <td>…</td>
              <td><a href="nyuryoku-csv.asp">CSVファイル転送</a></td>

            </tr>
            <tr> 
              <td>CSVファイル転送についての説明はここをクリック</td>
              <td>…</td>
              <td><a href="help07.asp">ヘルプ</a></td>
            </tr>
          </table>
              </form>
				</center>
</td></tr></table>


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
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
        If bError Then
		    ' コールサイン／次航入力
		    WriteLog fs, "3001","船社／ターミナル入力","10", strOption
		Else
		    ' コールサイン／次航入力
		    WriteLog fs, "3001","船社／ターミナル入力","00", ","
		End If
    Else
	    ' コールサイン／次航入力
	    WriteLog fs, "3001","船社／ターミナル入力","10", strOption

        ' 本船動静表示画面へリダイレクト
        Response.Redirect "nyuryoku-port.asp"    '本船動静表示画面
    End If
%>
