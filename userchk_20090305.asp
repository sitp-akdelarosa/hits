<%@Language="VBScript" %>
<%
'for each name in session.contents
'	response.write(name &"===="& session(name) &"<br>")
'next
'response.end


%>

<!--#include file="Common.inc"-->

<%


'特定画面へのリンク時にログを出力する
Sub CheckLinkLog
	Dim iNum,iWrkNum
    Select Case strLinkID
        Case "hits.asp"      strLinkNamne = "ストックヤード活用"
							iNum = "9002"
							iWrkNum = "00"
        Case "gate.asp"      strLinkNamne = "ゲート通行時間予約"
        Case Else            strLinkNamne = ""
    End Select
    If strLinkNamne<>"" Then
        ' File System Object の生成
        Set fs=Server.CreateObject("Scripting.FileSystemobject")

        ' リンク情報を出力
        WriteLog fs, iNum,strLinkNamne,iWrkNum, ","
    End If
End Sub

%>

<%


'画面の表示
Function DispLogIn(sError)
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function Check(){
  f=document.usercheck;
  userid = f.user.value;
  ret = CheckEisuji(userid);
  if(ret==false){
    alert("会社コードは半角英数字で入力してください。");
    return false;
  }
  return true;
}


function CheckEisuji(str){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/loginback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここからログイン入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.17 -->

		<table border=0><tr><td height=65></td></tr></table>

        <table border="0" cellspacing="4" cellpadding="0" bgcolor="#ff9933">
          <tr>
           <td>
		  <table border="0" cellspacing="3" cellpadding="4" bgcolor="#ffffff">
          <tr>
           <td>
        <table width="500" border="0" cellspacing="0" cellpadding="5" bgcolor="#FFFFCC">
          <tr>
           <td>
              <table width=100%>
                <tr>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>         <td><img src="gif/1.gif" width="1" height="1"></td>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
		</tr>
		<tr>
		 <td></td>		 
                  <td align="center">

      <table>
        <tr>
          <td nowrap align="center"> 
            <form name="usercheck" action="userchk.asp" method="put">
              <dl>
                <dd>会社コードとパスワードを入力し、『送信』ボタンをクリックしてください。 
              </dl>
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
                <tr> 
                  <td bgcolor="#ff8833" nowrap align=center valign=middle> <font color="#FFFFFF"><b>会社コード</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=100>
							<input type=text name=user value="" size=7 maxlength=5>
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
                  <td bgcolor="#ff8833" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>パスワード</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=100>
							<input type=password name=pass size=10 maxlength=8>
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
                <input type="submit" value=" 送　信 " onClick="return Check()"></center>
              </form>

          </td>
        </tr>
      </table>
<%
            If sError<>"" Then
                ' エラーメッセージの表示
                DispErrorMessage sError
            End If
%>
	  </td>
	  <td>
	  </td>
	 </tr>
        <tr>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
                  <td><img src="gif/1.gif" width=1 height=1></td>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
	          </td>
             </tr>
           </table>
	  	  </td>
        </tr>
      </table>
	  	  </td>
        </tr>
      </table>
	  	  </td>
        </tr>
      </table>
<br><br><br>
<a href="touroku/index.html" target="new">会社コード登録方法</a>
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
<!-------------ログイン画面終わり--------------------------->
<%
	If InStr(Request.QueryString("link"),"-expentry.asp")<>0 Then
		DispMenuBarBack "expentry.asp"
	ElseIf InStr(Request.QueryString("link"),"-impentry.asp")<>0 Then
		DispMenuBarBack "impentry.asp"
	Else
		DispMenuBarBack "index.asp"
	End If
%>
</body>
</html>

<%
End Function
%>

<%
'画面の表示
Function DispError
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
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.17 -->
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>ログイン</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>
      <table>
        <tr>
          <td nowrap align=center>
			<BR><BR>
            <dl>
				<img src="gif/error2.gif" width=210 height=63>
            </dl>
			<BR>
<%
            ' エラーメッセージの表示
            DispErrorMessage "ログインエラーのため、使用できません。"
%>
          </td>
        </tr>
      </table>
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
	If InStr(Request.QueryString("link"),"-expentry.asp")<>0 Then
		DispMenuBarBack "expentry.asp"
	ElseIf InStr(Request.QueryString("link"),"-impentry.asp")<>0 Then
		DispMenuBarBack "impentry.asp"
	Else
		DispMenuBarBack "index.asp"
	End If
%>
</body>
</html>

<%
End Function
%>

<%
' リンク画面を表示してよいかどうかのチェック
Function CheckLinkKind(iNum,iWrkNum)
    ' 戻り画面情報を取得
    strLinkID = Session.Contents("linkid")

    strError=""
    Select Case strLinkID
        Case "nyuryoku-in1.asp"             ' 船社／ターミナル入力
             If strUserKind<>"船社" And strUserKind<>"港運" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>船社、港運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
        Case "nyuryoku-kaika.asp", "nyuryoku-kaika2.asp"           ' 海貨入力  'Updated by seiko-denki 2003.07.21
             If strUserKind<>"海貨" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
        Case "nyuryoku-te.asp"              ' ターミナル入力
             If strUserKind<>"港運" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>港運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
        Case "rikuun1.asp"                  ' 陸運入力
             If strUserKind<>"陸運" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>陸運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
        Case "ms-kaika.asp"                 ' 松下仕様海貨入力
             If strUserKind<>"海貨" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
' Commented by seiko-denki 2003.07.07
'        Case "ms-expentry.asp?kind=1"       ' 松下仕様輸出コンテナ照会
'             If strUserKind<>"海貨" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
'        Case "ms-expentry.asp?kind=2"       ' 松下仕様輸出コンテナ照会
'             If strUserKind<>"陸運" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>陸運</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
'        Case "ms-expentry.asp?kind=3"       ' 松下仕様輸出コンテナ照会
'             If strUserKind<>"荷主" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>荷主</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
'        Case "ms-expentry.asp?kind=4"       ' 松下仕様輸出コンテナ照会
'             If strUserKind<>"港運" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>港運</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
'        Case "ms-impentry.asp?kind=1"       ' 松下仕様輸入コンテナ照会
'             If strUserKind<>"海貨" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
'        Case "ms-impentry.asp?kind=2"       ' 松下仕様輸入コンテナ照会
'             If strUserKind<>"陸運" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>陸運</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
'        Case "ms-impentry.asp?kind=3"       ' 松下仕様輸入コンテナ照会
'             If strUserKind<>"荷主" Then
'                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>荷主</font><font color=#008800>でのログイン時のみご使用になれます。"
'             End If
' End of Comment by seiko-denki 2003.07.07
' Added by seiko-denki 2003.07.07
        Case "ms-expentry.asp"       ' 松下仕様輸出コンテナ照会
             If strUserKind<>"海貨" And strUserKind<>"陸運" And strUserKind<>"荷主" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨、陸運、荷主</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
        Case "ms-impentry.asp"       ' 松下仕様輸出コンテナ照会
             If strUserKind<>"海貨" And strUserKind<>"陸運" And strUserKind<>"荷主" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨、陸運、荷主</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
' End of Addition by seiko-denki 2003.07.07
        Case "pickselect.asp"             ' 空コンピックアップシステム
             If strUserKind="船社" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨、陸運、荷主、港運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If

        Case "hits.asp"                     ' ストックヤード活用
        Case "gate.asp"                     ' ゲート通行時間予約

        Case "sokuji.asp"                   ' 即時搬出システム
             If strUserKind<>"海貨" And strUserKind<>"港運" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨、港運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
' Added by seiko-denki 2003.12.25
        Case "SendStatus/sst000F.asp"             ' ステータス配信
             If strUserKind="船社" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>海貨、陸運、荷主、港運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
' End of Addition by seiko-denki 2003.12.15



'''''Added 20040131
        Case "Shuttle/SYWB013.asp"                  ' シャトル予約
             If strUserKind<>"陸運" Then
                 strError="</font><font color=#008800>本機能は</font><font color=#ff0000>陸運</font><font color=#008800>でのログイン時のみご使用になれます。"
             End If
''''Added 20040131 END



    End Select

    If strError<>"" Then
        DispLogIn(strError)

        ' File System Object の生成
        Set fs=Server.CreateObject("Scripting.FileSystemobject")

        ' ログインエラー
        WriteLog fs, iNum,"ログインエラー",iWrkNum, strUserKind & "," & "入力内容の正誤:1(誤り)"
    End If
    CheckLinkKind = strError
End Function
%>

<%


    ' ログインエラー回数をチェック
    iError=CInt( Session.Contents("loginerror") )
    If iError>3 Then
        DispError
        Response.End
    End If

    ' 指定引数の取得（戻り画面情報）
    Dim strLinkID
    strLinkID = Request.QueryString("link")
    If strLinkID<>"" Then
        ' 戻り画面情報をセッション変数に設定
        Session.Contents("linkid") = strLinkID
        ' ログインエラー回数をクリア
        Session.Contents("loginerror") = 0
        iError=0
    End If

    ' 工事中の間、ユーザーＩＤチェックをしない画面
    Select Case strLinkID
        Case "hits.asp", "gate.asp"
            CheckLinkLog

            ' 戻り画面へリダイレクト
            Response.Redirect strLinkID
            Response.End
        Case Else
    End Select

    ' セッションの有効性をチェック
    Dim strSessionLink
    strSessionLink = Trim(Session.Contents("sessionlink"))
    ' セッションが無効なとき
    If strSessionLink="" Then
        ' セッション開始をセッション変数に設定
        Session.Contents("sessionlink") = "on"

        ' セッション切れが無効な画面のとき、メニューに戻る

    End If



	Dim iNum,iWrkNum
' ユーザーIDが必要な画面かどうかチェック
' Select Case strLinkID
' strLinkIDだとエラー時のログが取得できないのでセッションに変更しました	2002/2/21

		Select Case Session.Contents("linkid")
		' ユーザーIDが必要な画面
		Case ""
		Case "hits.asp", "gate.asp"
		Case "nyuryoku-in1.asp"
				iNum = 3000
				iWrkNum = 10
		Case "nyuryoku-kaika.asp", "nyuryoku-kaika2.asp"  'Updated by seiko-denki 2003.07.21
				iNum = 4000
				iWrkNum = 10
		Case "nyuryoku-te.asp"
				iNum = 5000
				iWrkNum = 10
		Case "rikuun1.asp"
				iNum = 6000
				iWrkNum = 10
'		Case "ms-expentry.asp?kind=1"   ' Commented by seiko-denki 2003.07.07
'				iNum = 1100
'				iWrkNum = 11
'		Case "ms-expentry.asp?kind=2"
'				iNum = 1100
'				iWrkNum = 12
'		Case "ms-expentry.asp?kind=3"
'				iNum = 1100
'				iWrkNum = 13
'		Case "ms-expentry.asp?kind=4"
'				iNum = 1100
'				iWrkNum = 14
'		Case "ms-impentry.asp?kind=1"
'				iNum = 2100
'				iWrkNum = 11
'		Case "ms-impentry.asp?kind=2"
'				iNum = 2100
'				iWrkNum = 12
'		Case "ms-impentry.asp?kind=3"
'				iNum = 2100
'				iWrkNum = 13  ' End of Comment by seiko-denki 2003.07.07
		Case "ms-expentry.asp"
				iNum = 1100
				iWrkNum = 11
		Case "ms-impentry.asp"
				iNum = 2100
				iWrkNum = 11
		Case "sokuji.asp"
				iNum = 7000
				iWrkNum = 10
		Case "pickselect.asp"
				iNum = "a100"
				iWrkNum = 10
		Case "predef/dmi000F.asp"
				iNum = "b000"
				iWrkNum = 10
		Case "SendStatus/sst000F.asp"  ' Added by seiko-denki 2003.12.25
				iNum = "c000"
				iWrkNum = 10             ' End of Addition by seiko-denki 2003.12.15
		Case "Shuttle/SYWB013.asp"		''''Added 20040131
				iNum = "d000"							''''Added 20040131
				iWrkNum = 10							''''Added 20040131
		' ユーザーIDが不要な画面
		Case "sokuji-kaika-list.asp", "sokuji-koun-list.asp"
		Case Else
				' 戻り画面へリダイレクト
 				CheckLinkLog
				Response.Redirect strLinkID
				Response.End
	End Select





    ' ユーザーIDの有効性をチェック
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))

    ' 指定引数の取得(ユーザーＩＤ)
    Dim strInputUserID, strInputPassWoed
    strInputUserID = UCase(Trim(Request.QueryString("user")))
    strInputPassWord = UCase(Trim(Request.QueryString("pass")))

    ' ユーザーIDが有効なとき
    If strUserID<>"" And strInputUserID="" Then
        ' ユーザ種類がマッチしているかチェックする
        strUserKind=Session.Contents("userkind")
        strError = CheckLinkKind(iNum,iWrkNum)
        If strError="" Then
            ' 戻り画面情報を取得
            strLinkID = Session.Contents("linkid")

            CheckLinkLog

            ' 戻り画面へリダイレクト
            Response.Redirect strLinkID
        Else
            ' ログインエラー回数をカウントアップ
            iError=iError+1
            Session.Contents("loginerror") = iError
        End If
    Else
        ' エラーフラグのクリア
        bOK = false
        bError = false

        If strInputUserID<>"" Then
            ' 入力ユーザーＩＤのチェック
            ConnectSvr conn, rsd
'=========== 03/07/17 変更 =================================================================
			sql="select FullName,UserType from mUsers"
			sql=sql&" where UserCode='" & strInputUserID & "' and PassWord='" & strInputPassWord & "'"
			'SQLを発行してユーザーＩＤを検索
			rsd.Open sql, conn, 0, 1, 1
			If Not rsd.EOF Then
				bOK = true
				' ログインＩＤをセッション変数に設定
				Session.Contents("userid") = strInputUserID
				' ログイン種別をセッション変数に設定
				Select Case Trim(rsd("UserType"))
					Case "1"
						Session.Contents("userkind") = "荷主"
					Case "2"
						Session.Contents("userkind") = "海貨"
					Case "3"
						Session.Contents("userkind") = "船社"
					Case "4"
						Session.Contents("userkind") = "港運"
					Case "5"
						Session.Contents("userkind") = "陸運"
				End Select
				' ログイン名をセッション変数に設定
				Session.Contents("username") = Trim(rsd("FullName"))
			End If
			rsd.Close
'=============================================================================================

'=========== 03/07/17 コメントアウト =================================================================
            ' 荷主コードチェック
'             sql = "SELECT FullName FROM mShipper WHERE Shipper='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
            'SQLを発行してユーザーＩＤを検索
'            rsd.Open sql, conn, 0, 1, 1
'            If Not rsd.EOF Then
'                bOK = true
                ' ログインＩＤをセッション変数に設定
'                Session.Contents("userid") = strInputUserID
                ' ログイン種別をセッション変数に設定
'                Session.Contents("userkind") = "荷主"
                ' ログイン名をセッション変数に設定
'                Session.Contents("username") = Trim(rsd("FullName"))
'            End If
'            rsd.Close

'            If bOK=false Then
                ' 海貨コードチェック
'                sql = "SELECT FullName FROM mForwarder WHERE Forwarder='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQLを発行してユーザーＩＤを検索
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ログインＩＤをセッション変数に設定
'                    Session.Contents("userid") = strInputUserID
                    ' ログイン種別をセッション変数に設定
'                    Session.Contents("userkind") = "海貨"
                    ' ログイン名をセッション変数に設定
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'            If bOK=false Then
                ' 陸運コードチェック
'                sql = "SELECT FullName FROM mTrucker WHERE Trucked='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQLを発行してユーザーＩＤを検索
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ログインＩＤをセッション変数に設定
'                    Session.Contents("userid") = strInputUserID
                    ' ログイン種別をセッション変数に設定
'                    Session.Contents("userkind") = "陸運"
                    ' ログイン名をセッション変数に設定
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'            If bOK=false Then
                ' 船社コードチェック
'                sql = "SELECT FullName FROM mShipLine WHERE ShipLine='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQLを発行してユーザーＩＤを検索
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ログインＩＤをセッション変数に設定
'                    Session.Contents("userid") = strInputUserID
                    ' ログイン種別をセッション変数に設定
'                    Session.Contents("userkind") = "船社"
                    ' ログイン名をセッション変数に設定
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'            If bOK=false Then
                ' 港運コードチェック
'                sql = "SELECT FullName FROM mOperator WHERE OpeCode='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQLを発行して港運マスターを検索
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ログインＩＤをセッション変数に設定
'                    Session.Contents("userid") = strInputUserID
                    ' ログイン種別をセッション変数に設定
'                    Session.Contents("userkind") = "港運"
                    ' ログイン名をセッション変数に設定
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'=============================================================================================
            If bOK=false Then
                ' ユーザーＩＤエラーのとき
                bError=true
                strError = "入力された内容に間違いがあります。"
                ' ログインエラー回数をカウントアップ
                iError=iError+1
                Session.Contents("loginerror") = iError
            End If

            conn.Close
        End If

        If Not bOK Then
            ' File System Object の生成
            Set fs=Server.CreateObject("Scripting.FileSystemobject")

            ' ログイン
            If strInputUserID<>"" Then
                WriteLog fs, iNum,"ログイン",iWrkNum, strInputUserID & "," & "入力内容の正誤:1(誤り)" & iError
            Else
                WriteLog fs, iNum,"ログイン", "00",","
            End If

            If iError>3 Then
                DispError
            Else
                If Not bError Then
                    strError=""
                    ' ログインエラー回数をカウントアップ
                    iError=iError+1
                    Session.Contents("loginerror") = iError
                End If
                DispLogIn(strError)
            End If
        Else
            ' ユーザ種類がマッチしているかチェックする
            strUserKind=Session.Contents("userkind")
            strError = CheckLinkKind(iNum,iWrkNum)
            If strError="" Then
                ' 戻り画面情報を取得
                strLinkID = Session.Contents("linkid")

                CheckLinkLog

                ' 戻り画面へリダイレクト
                Response.Redirect strLinkID
            Else
                ' ユーザ情報クリア
                    Session.Contents("userid") = ""
                    Session.Contents("userkind") = ""
                    Session.Contents("username") = ""
                ' ログインエラー回数をカウントアップ
                iError=iError+1
                Session.Contents("loginerror") = iError
            End If
        End If
    End If
%>
