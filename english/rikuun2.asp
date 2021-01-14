<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "rikunn1.asp"

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' DBの接続
    ConnectSvr conn, rs

    ' ユーザ種類を取得する
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"             'トップ
        Response.End
    End If

Dim vCtnoE, vCtnoS
Dim sCntNo
Dim sUserID
Dim sSQL
Dim sErrMsg
Dim sErrOpt

sErrMSg = ""
sErrOpt = ""

vCtnoE = Trim(Request.QueryString("cntnrnoe"))
vCtnoS = Trim(Request.QueryString("cntnrnos"))

If (IsEmpty(vCtnoE) Or vCtnoE = "") And (IsEmpty(vCtnoS) Or vCtnoS = "") Then
	sErrMsg = "コンテナ未入力"
End If

If sErrMsg = "" Then

	'該当するコンテナを探す
	If IsEmpty(vCtnoE) Or vCtnoE = "" Then
		'コンテナ番号の数値部分のみ入力されている場合
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container GROUP BY RTrim([ContNo]), ContNo "
		sSQL = sSQL & "HAVING (((RTrim([ContNo])) Like '%" & vCtnoS & "'))"
	Else
		'コンテナ番号の英字部分、数値部分ともに入力されている場合
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container "
		sSQL = sSQL & "WHERE RTrim([ContNo]) = '" & UCase(vCtnoE) & vCtnoS & "'"
	End If
	rs.Open sSQL, conn, 0, 1, 1
	If rs.Eof Then
		sErrMsg = "該当コンテナなし"
		sErrOpt = vCtnoS
	Else
		sCntNo = rs("CT")		'コンテナ番号再設定
		rs.MoveNext
		Do While Not rs.EOF
			sCntNo2 = rs("CT")
			rs.MoveNext
			If sCntNo<>sCntNo2 Then
				sErrMsg = "ｺﾝﾃﾅ複数存在"
				sErrOpt = vCtnoS
				Exit Do
			End If
		Loop
	End If
	rs.Close

    ' 陸運入力
	If sErrMsg = "" Then
        WriteLog fs, "6001", "陸運入力-コンテナ入力", "10", vCtnoE & "/" & vCtnoS & "," & "入力内容の正誤:0(正しい)"
		WriteLog fs, "6002", "陸運入力-完了時刻入力(Web)", "00", sCntNo & ","
    Else
        WriteLog fs, "6001", "陸運入力-コンテナ入力", "10", vCtnoE & "/" & vCtnoS & "," & "入力内容の正誤:1(誤り)" & sErrMsg
    End If

'	If sErrMsg = "" Then
'		' 今回検索したコンテナ番号をユーザテーブルに保存(次回にデフォルトで表示する為)
'		sSQL = "SELECT lUserTable.BeforeCntnrNo FROM lUserTable WHERE lUserTable.UserID='" & sUserID & "'"
'		rs.Open sSQL, conn, 2, 2
'		If Not rs.Eof Then
'			rs("BeforeCntnrNo") = sCntNo
'			rs.Update
'		End If
'		rs.Close
'	End If

'	conn.Close
End If
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
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="gif/rikuunt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
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
<%
    If sErrMsg<>"" Then
%>
<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>コンテナNo.入力</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
		<br><br>
<%
    DispErrorMessage sErrMsg
%>

<%
    Else
%>
<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>完了作業送信画面</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
		<br><br>
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td nowrap>
              完了した作業を選択して下さい。 <br>
			  『決定』をクリックすると現在の時間が入力されます。</td>
          </tr>
        </table>
        <br>
		          <form name=select action="rikuun3.asp">
			  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr> 
                    <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>コンテナNo.</b></font></td>
                    <td nowrap>
<%
    Response.Write sCntNo
    Session.Contents("cntnrno")=sCntNo
%>
                    </td>
                  </tr>
                  <tr> 
                    <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>完了した作業</b></font></td>
                    <td nowrap> 
                      <input type="radio" name="operation" value="C" checked>
					  （輸出）空倉庫着<br>
                      <input type="radio" name="operation" value="D">
                      （輸出）バンニング完了<br>
                      <input type="radio" name="operation" value="A">
                      （輸入）実入倉庫着<br>
                      <input type="radio" name="operation" value="B">
                      （輸入）デバン完了<br>
                    </td>
                  </tr>
                </table>
          <br>
          <input type=submit value="   決定   ">
</form>
<%
    End If
%>
		</center></td>
 </tr>
 <tr>
    <td valign="bottom">
<%
    DispMenuBar
%>
    </td>
  </tr>
</table>
 </td>
 </tr>
 </table>
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "rikuun1.asp"
%>
</body>
</html>