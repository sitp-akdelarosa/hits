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
    Dim iLine

    iLine = Trim(Request.QueryString("line"))

    ' 詳細表示行のデータの取得

    Dim iKensu	'表示件数(画面表示件数)
    Dim LineNo	'ファイルのラインカウンタ
    Dim iHitNo	'一致するファイル行数
    Dim sSensya, sSenmei, sJiko, sCallsign

    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
		If LineNo = 1 Then
	       iHitNo = clng(iLine) + 2
	'名称セット
           sSensya   = anyTmp(1)	'船社
           sSenmei   = anyTmp(3)	'船名
           If anyTmp(5) = anyTmp(6) Then	'次航
	       		sJiko = anyTmp(5)
		   Else
		        sJiko = anyTmp(5) & "/" & anyTmp(6)
		   End If
	       sCallsign = anyTmp(2)	'コールサイン

		End If

        If LineNo = iHitNo Then
           Exit Do
        End If
    Loop
    ti.Close

	' 着岸予定
    If  anyTmp(2) = ""  or anyTmp(2) = vbNull  Then
    Else	
	ayearval = Left(anyTmp(2), 4)
	amonthval = Mid(anyTmp(2), 6, 2)
	adayval = Mid(anyTmp(2), 9, 2)
	ahourval = Mid(anyTmp(2), 12, 2)
	aminval = Mid(anyTmp(2), 15, 2)
    End If

	' 着岸完了
    if  anyTmp(3) = ""  or anyTmp(3) = vbNull  Then
    Else	
	tyearval = Left(anyTmp(3), 4)
	tmonthval = Mid(anyTmp(3), 6, 2)
	tdayval = Mid(anyTmp(3), 9, 2)
	thourval = Mid(anyTmp(3), 12, 2)
	tminval = Mid(anyTmp(3), 15, 2)
    End If

	' 離岸完了
    if  anyTmp(5) = ""  or anyTmp(5) = vbNull  Then
    Else	
	dyearval = Left(anyTmp(5), 4)
	dmonthval = Mid(anyTmp(5), 6, 2)
	ddayval = Mid(anyTmp(5), 9, 2)
	dhourval = Mid(anyTmp(5), 12, 2)
	dminval = Mid(anyTmp(5), 15, 2)
    End If

	' 着岸 Long Schedule
    if  anyTmp(6) = ""  or anyTmp(6) = vbNull  Then
    Else	
	cyearval = Left(anyTmp(6), 4)
	cmonthval = Mid(anyTmp(6), 6, 2)
	cdayval = Mid(anyTmp(6), 9, 2)
'	chourval = Mid(anyTmp(6), 12, 2)
'	cminval = Mid(anyTmp(6), 15, 2)
    End If

	' 離岸 Long Schedule
    if  anyTmp(7) = ""  or anyTmp(7) = vbNull  Then
    Else	
	ryearval = Left(anyTmp(7), 4)
	rmonthval = Mid(anyTmp(7), 6, 2)
	rdayval = Mid(anyTmp(7), 9, 2)
'	rhourval = Mid(anyTmp(7), 12, 2)
'	rminval = Mid(anyTmp(7), 15, 2)
    End If

    ' 本船動静入力表示画面
    WriteLog fs, "3004","船社／ターミナル入力-本船動静入力", "02", anyTmp(1) & "," 
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {
	if (ChkSend("着岸予定時刻",
				document.con.ayear.value, 
				document.con.amonth.value,
				document.con.aday.value,
				document.con.ahour.value,
				document.con.amin.value) && 
           	ChkSend("着岸完了時刻",
				document.con.tyear.value, 
				document.con.tmonth.value,
				document.con.tday.value,
				document.con.thour.value,
				document.con.tmin.value) &&
           	ChkSend("離岸完了時刻",
				document.con.dyear.value, 
				document.con.dmonth.value,
				document.con.dday.value,
				document.con.dhour.value,
				document.con.dmin.value) &&
           	ChkSend("着岸 Long Schedule",
				document.con.cyear.value, 
				document.con.cmonth.value,
				document.con.cday.value,
//				document.con.chour.value,
//				document.con.cmin.value) &&
				"","") &&
	   		ChkSend("離岸 Long Schedule", 
				document.con.ryear.value, 
				document.con.rmonth.value,
				document.con.rday.value,
//				document.con.rhour.value,
//				document.con.rmin.value)) {
				"","") ) {
		return true;
	}
	return false;
}

function ChkSend(Name, sYear, sMonth, sDay, sHour, sTime) {

	if (Name == "着岸予定時刻") {
		if (sYear == "" ||  sMonth == "" || sDay == "") {
			window.alert(Name + "は必須入力です。");
			return false;
		}
	}
	else {
		if (sYear == "" &&  sMonth == "" && sDay == "" &&  sHour == ""  && sTime == "") {
			return true;
		}
	}
	
	if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* 年のチェック */
			window.alert(Name + "の年の入力が不正です。");
			return false;
	}
	if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* 月のチェック */
			window.alert(Name + "の月の入力が不正です。");
			return false;
	}
	if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {		/* 日のチェック */
			window.alert(Name + "の日の入力が不正です。");
			return false;
	}

	if (!(sHour > 0 || sHour <= 0)|| sHour < 0 || sHour > 24  ) {		/* 時のチェック */
			window.alert(Name + "の時の入力が不正です。");
			return false;
	}

	if (!(sTime > 0 || sTime <= 0)|| sTime < 0 || sTime > 59  ) {		/* 分のチェック */
			window.alert(Name + "の分の入力が不正です。");
			return false;
	}

	if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
	   (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
			window.alert(Name + "の日の入力が不正です。");
			return false;
	}
	return true;
}
/* 削除押下時の処理 */
function Click_Del() {
		location.href = "nyuryoku-sch-del.asp?line=" + document.con.iLine.value
		return true;
}

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから一覧画面--------------------------->
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
	  <table border="0">
		<tr> 
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>本船動静入力　</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
   	  <br>     
	日付及び時間は、半角数字で入力して下さい。
	&nbsp;&nbsp;&nbsp;（例 ） 2002年2月25日 15時30分
<BR><BR>
      <table border=0><tr><td>
          <table border=1 cellpadding="3" cellspacing="1">
                <tr> 
                  <td bgcolor="#000099" backgrond="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>船社</b></font></td>
                  <td bgcolor="#FFFFFF" nowrap>
<%
    ' 船社名の表示
    Response.Write sSensya
%>
                  </td>
                  <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>船名</b></font></td>
                  <td bgcolor="#FFFFFF" nowrap>
<%
    ' 船名の表示
    Response.Write sSenmei
%>
                  </td>
                </tr>
          </table>
          <table border=1 cellpadding="3" cellspacing="1">
                <tr>
                  <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>Voyage No.</b></font></td>
                  <td bgcolor="#FFFFFF" nowrap>
<%
    ' 次航の表示
    Response.Write sJiko
%>
                  </td>
                  <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コールサイン</b></font></td>
                  <td bgcolor="#FFFFFF" nowrap>
<%
    ' コールサインの表示
    Response.Write sCallsign
%>
                  </td>
                </tr>
          </table>
          <br>
          <FORM NAME="con" METHOD="post" action="nyuryoku-sch-upd.asp?line=<%=iLine%>" onSubmit="return ClickSend()">
			<table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
                <tr>
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>港名</b></font>
                  </td>
                  <td nowrap bgcolor="#FFFFFF">
<%
    ' 港名称の表示
    Response.Write anyTmp(1)
%>
					<input type=hidden name=sportname value=<%=anyTmp(1)%>>
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>着岸予定時刻</b></font>
                  </td>
                  <td nowrap>
                    <input type=text name=ayear size=4 value="<%=ayearval%>" maxlength="4">年
                    <input type=text name=amonth size=2 value="<%=amonthval%>" maxlength="2">月
                    <input type=text name=aday size=2 value="<%=adayval%>" maxlength="2">日　
                    <input type=text name=ahour size=2 value="<%=ahourval%>" maxlength="2">時
                    <input type=text name=amin size=2 value="<%=aminval%>" maxlength="2">分
					&nbsp;&nbsp;<font size=1 color="#ee2200">[ 必須入力 ]</font>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>着岸完了時刻</b></font>
                  </td>
                  <td nowrap>
                    <input type=text name=tyear size=4 value="<%=tyearval%>" maxlength="4">年
                    <input type=text name=tmonth size=2 value="<%=tmonthval%>" maxlength="2">月
                    <input type=text name=tday size=2 value="<%=tdayval%>" maxlength="2">日　
                    <input type=text name=thour size=2 value="<%=thourval%>" maxlength="2">時
                    <input type=text name=tmin size=2 value="<%=tminval%>" maxlength="2">分
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>離岸完了時刻</b></font>
                  </td>
                  <td nowrap>
                    <input type=text name=dyear size=4 value="<%=dyearval%>" maxlength="4">年
                    <input type=text name=dmonth size=2 value="<%=dmonthval%>" maxlength="2">月
                    <input type=text name=dday size=2 value="<%=ddayval%>" maxlength="2">日　
                    <input type=text name=dhour size=2 value="<%=dhourval%>" maxlength="2">時
                    <input type=text name=dmin size=2 value="<%=dminval%>" maxlength="2">分
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>着岸 Long Schedule</b></font>
                  </td>
                  <td nowrap>
                    <input type=text name=cyear size=4 value="<%=cyearval%>" maxlength="4">年
                    <input type=text name=cmonth size=2 value="<%=cmonthval%>" maxlength="2">月
                    <input type=text name=cday size=2 value="<%=cdayval%>" maxlength="2">日　
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>離岸 Long Schedule</b></font>
                  </td>
                  <td nowrap>
                    <input type=text name=ryear size=4 value="<%=ryearval%>" maxlength="4">年
                    <input type=text name=rmonth size=2 value="<%=rmonthval%>" maxlength="2">月
                    <input type=text name=rday size=2 value="<%=rdayval%>" maxlength="2">日　
                  </td>
                </tr>
              	<input type=hidden name=iLine VALUE="<%=iLine%>">
            </table>
            <br><br>
            <center>
                <input type=submit value=" 入  力 " name="nyuryoku">
                <input type="button" value=" 削　除 " name="Del" onclick="Click_Del()">
                <input type="button" value=" キャンセル" onclick="history.back()">
            </center>
          </form>
      </table>
      </center>
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
    DispMenuBarBack "nyuryoku-port.asp"
%>
</body>
</html>
