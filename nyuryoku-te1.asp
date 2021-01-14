<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-te.asp"

    ' エラーフラグのクリア
    bError = false

    ' 入力フラグのクリア
    bInput = true

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' セッション変数から港運コードを取得
    Dim strOpeCode
    strOpeCode = Trim(Session.Contents("userid"))
    strChoice = Trim(Session.Contents("choice"))

    ' 指定引数の取得
    Dim strCallSign
    Dim strVoyage
    Dim strBLNo
    Dim strYear
    Dim strMonth
    Dim strDay
    Dim strHour
    Dim strMin
    strCallSign = UCase(Trim(Request.QueryString("callsign")))
    strVoyage = UCase(Trim(Request.QueryString("voyage")))
    strBLNo = UCase(Trim(Request.QueryString("blno")))
    strYear = Trim(Request.QueryString("year"))
    strMonth = Trim(Request.QueryString("month"))
    strDay = Trim(Request.QueryString("day"))
    strHour = Trim(Request.QueryString("hour"))
    strMin = Trim(Request.QueryString("min"))
    If strChoice="bl" Then
	    strInput = strCallSign & "/" & strVoyage & "/" & strBLNo & "/" & strYear & "/" & strMonth & "/" & strDay & " " & strHour & ":" & strMin
	Else
	    strInput = strCallSign & "/" & strVoyage & "/" & strYear & "/" & strMonth & "/" & strDay & " " & strHour & ":" & strMin
	End If

    If strCallSign="" Or strVoyage="" Or strYear="" Or strMonth="" Or strDay="" Then
        If strCallSign<>"" Or strVoyage<>"" Or strYear<>"" Or strMonth<>"" Or strDay<>"" Or strHour<>"" Or strMin<>"" Then
            ' 入力が一部だけのとき エラーメッセージを表示
            bError = true
            strError = "入力が間違っています。"
            strOption = strInput & "," & "入力内容の正誤:1(誤り)"
		ElseIf strChoice="bl" And strBLNo<>"" Then
            ' 入力が一部だけのとき エラーメッセージを表示
            bError = true
            strError = "入力が間違っています。"
            strOption = strInput & "," & "入力内容の正誤:1(誤り)"
        Else
            bInput = false
        End If
    End If

    If bInput And Not bError Then
        ' 入力コールサインのチェック
        ConnectSvr conn, rsd
        sql = "SELECT FullName FROM mVessel WHERE VslCode='" & strCallSign & "'"
        'SQLを発行して船名マスターを検索
        rsd.Open sql, conn, 0, 1, 1
        If Not rsd.EOF Then
            strVesselName = Trim(rsd("FullName"))
            strOption = strInput & "," & "入力内容の正誤:0(正しい)"
        Else
            ' 該当レコードのないとき エラーメッセージを表示
            bError = true
            strError = "コールサインが間違っています。"
            strOption = strInput & "," & "入力内容の正誤:1(誤り)"
        End If
        rsd.Close
        If Not bError Then
            ' SQLを発行して本船動静を検索
            sql = "SELECT VoyCtrl FROM VslSchedule " & _
                  "WHERE VslCode='" & strCallSign & "' And DsVoyage='" & strVoyage & "'"
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                iVoyCtrl = rsd("VoyCtrl")
            Else
                ' 該当レコードのないとき エラーメッセージを表示
                bError = true
                strError = "Voyage No.が間違っています。"
                strOption = strInput & "," & "入力内容の正誤:1(誤り)"
            End If
            rsd.Close
        End If
        If Not bError And strChoice="bl" Then
            ' SQLを発行して輸入BLを検索
            sql = "SELECT ShipLine FROM BL " & _
                  "WHERE VslCode='" & strCallSign & "' And VoyCtrl=" & iVoyCtrl & " And BLNo='" & strBLNo & "'"
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                strShipLine = Trim(rsd("ShipLine"))
            Else
                ' 該当レコードのないとき エラーメッセージを表示
                bError = true
                strError = "BL番号が間違っています。"
                strOption = strInput & "," & "入力内容の正誤:1(誤り)"
            End If
            rsd.Close
        End If
        If Not bError Then
            ' 入力処理予定日をチェック

            ' 入力データを送信ファイルに出力
            strTmp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
                     Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
            strSeqNo = GetDailyTransNo()
            strFileName = Mid(strTmp,5,4) & strSeqNo & ".snd"

            strFileName="./send/" & strFileName
            ' テンポラリファイルのOpen
            Set ti = fs.OpenTextFile(Server.MapPath(strFileName),2,True)

			strMonth = DateFormat(strMonth)
			strDay = DateFormat(strDay)
			strHour = DateFormat(strHour)
            If strHour="" Then
                strHour="23"
            End If
			strMin = DateFormat(strMin)
            If strMin="" Then
                strMin="59"
            End If
			If strChoice="bl" Then
	            ti.WriteLine strSeqNo & ",IM15,R," & strTmp & ",Web - " & Session.Contents("userid") & ",," & _
	                         strCallSign & "," & strVoyage & "," & strBLNo & "," & strYear & strMonth & strDay & strHour & strMin
			Else
	            ti.WriteLine strSeqNo & ",IM15,R," & strTmp & ",Web - " & Session.Contents("userid") & ",," & _
	                         strCallSign & "," & strVoyage & ",," & strYear & strMonth & strDay & strHour & strMin
			End If

            ti.Close

            bError = true
            strError = "正常に更新されました。"
        End If
        conn.Close
    End If

    ' 搬入確認予定時刻入力(BL単位)
    If strChoice="bl" Then
		If bInput Then
	        WriteLog fs, "5002", "ターミナル入力-搬入確認予定時刻入力(BL単位)", "10", strOption
		Else
	        WriteLog fs, "5002", "ターミナル入力-搬入確認予定時刻入力(BL単位)", "00", ","
		End If
    Else
		If bInput Then
	        WriteLog fs, "5004", "ターミナル入力-搬入確認予定時刻入力(本船単位)", "10", strOption
		Else
	        WriteLog fs, "5004", "ターミナル入力-搬入確認予定時刻入力(本船単位)", "00", ","
		End If
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

	function ParamClear(){
		window.document.con.callsign.value  = "";
		window.document.con.voyage.value 	= "";
		window.document.con.year.value 		= "";
		window.document.con.month.value 	= "";
		window.document.con.day.value 		= "";
		window.document.con.hour.value 		= "";
		window.document.con.min.value 		= "";
<% If strChoice="bl" Then %>
		window.document.con.blno.value 		= "";
<%  End If %>
	}

	function DateCheck(){
		bErr  = true;
		mYear =	window.document.con.year.value;
		mMon  =	window.document.con.month.value;
		mDay  =	window.document.con.day.value;
		mHour = window.document.con.hour.value;
		mMin  =	window.document.con.min.value;

		if(!(mYear > 0 || mYear <= 0)|| mYear > 2100 || mYear < 1990){
            sName="年";
            bErr = false;
        }
		if(!(mMon > 0 || mMon <= 0)|| mMon  > 12   || mMon  < 1){
            sName="月";
            bErr = false;
        }
		if(!(mDay > 0 || mDay <= 0)|| mDay  > 31   || mDay  < 1){
            sName="日";
            bErr = false;
        }
		if(!(mHour > 0 || mHour <= 0)|| mHour > 23   || mHour < 0){
            sName="時";
            bErr = false;
        }
		if(!(mMin > 0 || mMin <= 0)|| mMin  > 59   || mMin  < 0){
            sName="分";
            bErr = false;
        }

        if (mDay>30+((mMon==4||mMon==6||mMon==9||mMon==11)?0:1) || 
           (mMon==2 && mDay>28+(((mYear%4==0 && mYear%100!=0) || mYear%400==0)?1:0)) ){
            sName="日";
            bErr = false;
		}

		if(!bErr) window.alert("処理予定日時の" + sName + "の入力が不正です。");
		return bErr;
	}

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから搬入確認予定時刻入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/terminal2t.gif" width="506" height="73"></td>
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
<table border=0><tr><td>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>
<% If strChoice="bl" Then %>
	個別搬入確認予定時刻入力( BL単位 )
<% Else %>
	一括搬入確認予定時刻入力( 本船単位 )
<%  End If %>
			</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
<center>
      <table>
        <tr>
          <td nowrap align=left>下記の項目を入力の上、『送信』ボタンをクリックして下さい。<BR>
			容器通関の時刻を入力します。
		  </td>
        </tr>
      </table>
      <FORM NAME="con" action="nyuryoku-te1.asp" onSubmit="return DateCheck()">
        <table border=0 cellpadding=0>
          <tr>
            <td align="center"> 
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF"> 
                    コールサイン</font></b></td>
                  <td>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=250>
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
						<td width=250>
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
                <tr>
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">搬入確認予定日時
                  </font></b></td>
                  <td>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=250>
<% 
	If strYear="" Then
		strYear = Year(Now)
	End If
	If strMonth="" Then
		strMonth = Month(Now)
	End If

If strChoice<>"bl" Then
	If strDay="" Then
		strDay = Day(Now)
	End If
End If

	If strHour="" Then
		strHour = "00"
	End If
	If strMin="" Then
		strMin = "00"
	End If
%>
		                    <input type=text name=year value="<%=strYear%>" size=4 maxlength=4>
		                    年
		                    <input type=text name=month value="<%=strMonth%>" size=2 maxlength=2>
		                    月
		                    <input type=text name=day value="<%=strDay%>" size=2 maxlength=2>
		                    日
		                    <input type=text name=hour value="<%=strHour%>" size=2 maxlength=2>
		                    時
		                    <input type=text name=min value="<%=strMin%>" size=2 maxlength=2>
		                    分
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
							<font size=1 color="#2288ff">[ 半角数値 ]</font>
						</td>
					  </tr>
					</table>
					&nbsp;&nbsp;&nbsp;<font size=-1>（例） 2002年2月25日 15時30分</font>
                  </td>
                </tr>
<% If strChoice="bl" Then %>
                <tr>
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">BL 
                    No.</font></b></td>
                  <td>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=250>
							<input type=text name=blno value="<%=strBLNo%>" size=20 maxlength=20>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>
<%  End If %>
              </table>
              <br>
              <INPUT TYPE=submit VALUE="　送信　">
              <INPUT TYPE=button VALUE="　クリア　" onClick="ParamClear()">
            </td>
          </tr>
         </table>
<%
    ' エラーメッセージの表示
    If bError Then
        If strError="正常に更新されました。" Then
            DispInformationMessage strError
        Else
            DispErrorMessage strError
        End If
    End If
%>

</center>
         <BR><BR>
         <table>
           <tr> 
             <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
             <td nowrap><b>ファイル転送</b></td>
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
              <td><a href="nyuryoku-tmnl-csv.asp">CSVファイル転送</a></td>
            </tr>
            <tr> 
              <td>CSVファイル転送についての説明はここをクリック</td>
              <td>…</td>
<% If strChoice="bl" Then %>
              <td><a href="help11.asp">ヘルプ</a></td>
<%  Else  %>
              <td><a href="help12.asp">ヘルプ</a></td>
<%  End If  %>
              
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
	DispMenuBarBack "nyuryoku-te.asp"
%>
</body>
</html>

<%
    Else
        ' 本船動静表示画面へリダイレクト
        Response.Redirect "nyuryoku-te.asp"    '本船動静表示画面
    End If
%>
