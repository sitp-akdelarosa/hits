<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-te.asp"

    ' 入力フラグのクリア
    bInput = true

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 指定引数の取得
    Dim strChoice
    strChoice = Request.QueryString("choice")
    If strChoice<>"" Then
        bInput = false
    End If

    ' セッション変数から港運コードを取得
    strOperator = Trim(Session.Contents("userid"))

    If bInput Then
        ' 搬入確認予定時刻入力
        WriteLog fs, "5001", "ターミナル入力", "00", ","
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
<!-------------ここから港運登録画面--------------------------->
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
      <table border=0>
        <tr>
          <td align=left>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                  <td nowrap><b>搬入確認予定時刻入力</b></td>
                  <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <FORM NAME="con" action="nyuryoku-te.asp">
            <br>
            <center>
              搬入確認予定時刻を入力します。<br>
              次のいずれかの方法を選択して『送信』ボタンをクリックしてください。<br>
              <br>
              <table border="0" cellspacing="2" cellpadding="3">
                <tr>
                  <td>
                    <input type="radio" name="choice" value="bl" checked>個別(BL単位)
                  </td>
                </tr>
                <tr>
                  <td>
                    <input type="radio" name="choice" value="vsl">一括(本船単位)
                  </td>
                </tr>
              </table>
              <br>
              <br>
                <input type=submit value=" 送　信 "><br><br><br>
              </center>
            </form>
          </td>
        </tr>
      </table>
      <BR>
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
<!-------------港運登録画面終わり--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>

</body>
</html>

<%
    Else
        Session.Contents("choice")=strChoice
        ' 搬入確認予定時刻入力画面へリダイレクト
        Response.Redirect "nyuryoku-te1.asp"    '搬入確認予定時刻入力画面
    End If
%>
