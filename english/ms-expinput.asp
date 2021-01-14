<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' 指定引数の取得
    Dim strKind       '入力種類(1=届時刻,2=完了時刻)
    Dim iLine         '入力行
    Dim strRequest    '戻り先
    strKind=Trim(Request.QueryString("kind"))
    iLine=CInt(Trim(Request.QueryString("line")))
    strRequest=Trim(Request.QueryString("request"))

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "index.asp"             'メニュー画面へ
        Response.End
    End If
    strFileName="./temp/" & strFileName

	Dim iNum
    ' 輸出陸運情報入力
    If strKind="1" Then
		iNum = ""
       strTitle="(輸出)空コンテナ倉庫到着時刻"
    Else
		iNum = "1107"
       strTitle="(輸出)バンニング完了時刻"
    End If
    WriteLog fs, iNum,"輸出コンテナ照会-バンニング完了時刻入力","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 詳細表示行のデータの取得
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLine=LineNo Then
           Exit Do
        End If
    Loop
    ti.Close

    Session.Contents("editkind")=strKind         ' 入力種類を記憶
    Session.Contents("editline")=iLine           ' 編集行を記憶
    Session.Contents("request")=strRequest       ' 戻り画面を記憶
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
    If strKind="1" Then
%>
function ClickSend() {
    return (ChkSend("届け時刻",
                    document.con.Year.value, 
                    document.con.Month.value, 
                    document.con.Day.value, 
                    document.con.Hour.value, 
                    document.con.Min.value));
}
<%
    Else
%>
function ClickSend() {
    return (ChkSend("完了時刻",
                    document.con.Year.value, 
                    document.con.Month.value, 
                    document.con.Day.value, 
                    document.con.Hour.value, 
                    document.con.Min.value));
}
<%
    End If
%>
// 入力チェック
function ChkSend(sMes, sYear, sMonth, sDay, sHour, sMin ) {
    if (sYear == "" ||  sMonth == "" || sDay == "" || sHour == "" || sMin == "") {
        window.alert(sMes+"が未入力です。");
        return false;
    }
    if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* 年のチェック */
        window.alert(sMes+"の年の入力が不正です。");
        return false;
    }
    if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* 月のチェック */
        window.alert(sMes+"の月の入力が不正です。");
        return false;
    }
    if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {		/* 日のチェック */
        window.alert(sMes+"の日の入力が不正です。");
        return false;
    }
    if (!(sHour > 0 || sHour <= 0)|| sHour < 0 || sHour > 24  ) {		/* 時のチェック */
        window.alert(sMes+"の時の入力が不正です。");
        return false;
    }
    if (!(sMin > 0 || sMin <= 0)|| sMin < 0 || sMin > 59  ) {		/* 分のチェック */
        window.alert(sMes+"の分の入力が不正です。");
        return false;
    }
    if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
       (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
        window.alert(sMes+"の日の入力が不正です。");
        return false;
    }
    return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/exprikuun.gif" width="506" height="73"></td>
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
      <table>
        <tr>
          <td>下記の項目を入力の上、送信ボタンをクリックして下さい。</td>
        </tr>
      </table>
      <FORM NAME="con" METHOD="post" action="ms-expinput-syori.asp" onSubmit="return ClickSend()">
		<input type=hidden name=title value="<%=strTitle%>">
        <table border=0 cellpadding=0 bordercolor="#999999">
          <tr> 
            <td align="center"> 
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                    荷主名称</font></b></td>
                  <td bgcolor="#FFFFFF"> 
<% ' 荷主情報 - 名称
    Response.Write anyTmp(7)
%>
                  </td>
                </tr>
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                     荷主管理番号</font></b></td>
                  <td bgcolor="#FFFFFF"> 
<% ' 荷主情報 - 管理番号
    Response.Write anyTmp(14)
%>
                  </td>
                </tr>
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
<%
    If strKind="1" Then
        Response.Write "届け時刻"
    Else
        Response.Write "完了時刻"
    End If
%>
                    </font></b></td>
                  <td> 
<%
    If strKind="1" Then
        strTemp=anyTmp(47)
    Else
        strTemp=anyTmp(48)
    End If
    If strTemp="" Then
        strTemp=DispDateTime(Now,0)
    End If
    Response.Write "<input type=text name='Year' value='" & Left(strTemp,4) & "' size=4 maxlength='4'>年"
    Response.Write "<input type=text name='Month' value='" & Mid(strTemp,6,2) & "' size=2 maxlength='2'>月"
    Response.Write "<input type=text name='Day' value='" & Mid(strTemp,9,2)  & "' size=2 maxlength='2'>日　"
    Response.Write "<input type=text name='Hour' value='" & Mid(strTemp,12,2)  & "' size=2 maxlength='2'>時"
    Response.Write "<input type=text name='Min' value='" & Mid(strTemp,15,2)  & "' size=2 maxlength='2'>分"
%>
					<font size=1 color="#2288ff">[半角数値]</font><BR>
					&nbsp;&nbsp;&nbsp;<font size=-1>（例） 2002年 2月 25日 15時 30分</font>
                  </td>
                </tr>
              </table>
              <br>
              <input type=submit value="　送信　">
              <input type="button" value="　中止　" onclick="history.back()">
            </td>
          </tr>
        </table>
      </form>
      <br>
      <br>
      <br>
      <br>
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
    If strRequest="ms-expdetail.asp" Then
        strTemp=strRequest & "?line=" & iLine
    Else
        strTemp=strRequest
    End If
    DispMenuBarBack strTemp
%>
</body>
</html>
