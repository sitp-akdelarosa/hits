<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'
	'	【輸入コンテナ情報入力】	新規入力、ＣＳＶ転送、更新選択画面
	'
%>

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"
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
<!-------------ここからログイン入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=95%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika6t.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>新規入力</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>

	<center>
      <table>
        <tr>
          <td nowrap align=center>
			輸入コンテナ情報を新規で入力する場合はここをクリック ･･･ <a href="ms-kaika-impcontinfo-new.asp?kind=1">新規入力</a>
		  </td>
		</tr>
	  </table>
	</center>

      <br><BR>
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
            <p>輸入コンテナ情報をファイル転送する場合はここをクリック</p>
          </td>
          <td>…</td>
          <td><a href="ms-kaika-impcontinfo-csv.asp">CSVファイル転送</a></td>
        </tr>
        <tr> 
          <td>CSVファイル転送についての説明はここをクリック</td>
          <td>…</td>
          <td><a href="help21.asp">ヘルプ</a></td>
        </tr>
      </table>
	</center>

	<BR><BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>対象データの指定</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>

	<center>
      <table>
        <tr>
          <td nowrap>
            <form method=post action="ms-kaika-impcontinfo-updatecheck.asp">
              <center>
				<table border=0 cellpadding=0>
				  <tr>
					<td align=left>
				絞り込みを行う場合は、下記フォームに適当な値を入力してから<BR>更新対象一覧ボタンを押して下さい。
				<BR><BR>条件を入力しないで照会を実行すると、全件表示されます。
					</td>
				  </tr>
				</table>
				<BR>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle> <font color="#FFFFFF"><b>荷主コード</b></font></td>
                  <td nowrap>
                    <input type=text name=suser size=7 maxlength=5> <font size=1 color="#2288ff">[ 半角英数 ]</font>
                  </td>
                </tr>
              </table>
              <br>
                <input type=submit value="更新対象一覧"> 
				</center>
              </center>
            </form>
		  </td>
		</tr>
	  </table>
	</center>


<%
            If bError Then
                ' エラーメッセージの表示
                DispErrorMessage strError
            End If
%>
          </td>
        </tr>
      </table>
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
<!-------------ログイン画面終わり--------------------------->
<%
    DispMenuBarBack "nyuryoku-kaika.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    WriteLog fs, "4109","海貨入力輸入コンテナ情報","00", ","
%>
