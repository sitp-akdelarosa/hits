<!--#include file="Common.inc"-->
<%
'for each name in session.contents
'	response.write(name &"===="& session(name) &"<br>")
'next
'response.end
%>
<%
	' 戻り画面
	dim ReturnUrl
	ReturnUrl = request.querystring("ReturnUrl")
	' 指定引数の取得(ユーザーＩＤ)
	Dim strInputUserID, strInputPassWord
	
	If UCase(Trim(Request.form("user")))<>"" and UCase(Trim(Request.form("pass")))<>"" then
		Session.Contents("userkind")=""
	End If

	If Session.Contents("userid")<>"" and Session.Contents("userkind") = "陸運" Then
		strInputUserID = Session.Contents("userid")
		ReturnUrl = ReturnUrl & "?UserId=" & strInputUserID
		response.redirect ReturnUrl
	Else
	strInputUserID = UCase(Trim(Request.form("user")))
	strInputPassWord = UCase(Trim(Request.form("pass")))
	End If

	bOK = false
	bError = false

	If strInputUserID<>"" Then
		' 入力ユーザーＩＤのチェック
		ConnectSvr conn, rsd

		' 陸運コードチェック
			sql="select FullName from mUsers"
			sql=sql&" where UserCode='" & strInputUserID & "' and PassWord='" & strInputPassWord & "' and UserType='5'"
		'SQLを発行してユーザーＩＤを検索
		rsd.Open sql, conn, 0, 1, 1
		If Not rsd.EOF Then
			bOK = true
			' ログインＩＤをセッション変数に設定
			Session.Contents("userid") = strInputUserID
			' ログイン種別をセッション変数に設定
			Session.Contents("userkind") = "陸運"
			' ログイン名をセッション変数に設定
			Session.Contents("username") = Trim(rsd("FullName"))
		End If
		rsd.Close

		If bOK=false Then
		    ' ユーザーＩＤエラーのとき
		    bError=true
		    strError = "入力された内容に間違いがあります。"
		    ' ログインエラー回数をカウントアップ
	'	    iError=iError+1
	'	    Session.Contents("loginerror") = iError
		End If

		conn.Close
	End If
	
	If bOK=true Then
		ReturnUrl = ReturnUrl & "?UserId=" & strInputUserID
		response.redirect ReturnUrl
	End If

%>


<html>
<head>
<title>ログイン</title>
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
	DisplayCodeListButton
%>
          </td>
        </tr>
      </table>
      <center>

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
            <form action="userchk2.asp?ReturnUrl=<%=ReturnUrl%>" method="post" name="usercheck">
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
                <input type=submit value=" 送　信 " onClick="return Check()"></center>
              </form>
          </td>
        </tr>
      </table>
	<% If strError<>"" then %>
		<table border=1 cellpadding="2" cellspacing="1">
		<tr>
			<td bgcolor="#FFFFFF">
				<table border="0">
				<tr>
					<td valign="middle"><img src="gif/error.gif"></td>
					<td><b><font color="red"><% =strError %></font></b></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	<% End If %>
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
</body>
</html>

