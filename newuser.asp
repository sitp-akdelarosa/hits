<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションの有効性をチェック
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))

    ' セッションが有効なとき
    If strUserID<>"" Then
        ' 戻り画面情報を取得
        strLinkID = Session.Contents("linkid")

        ' 戻り画面へリダイレクト
        Response.Redirect strLinkID
    Else
        ' エラーフラグのクリア
        bOK = false
        bError = false

        ' 指定引数の取得(会社名,メールアドレス)
        Dim strCompany
        Dim strMailAddress
        strCompany = Trim(Request.QueryString("campany"))
        strMailAddress = Trim(Request.QueryString("mail"))

        If strCompany<>"" Then
            ' ユーザーＩＤの最大チェック
            ConnectSvr conn, rsd

            sql = "SELECT UserID, CompanyName, MailAddress FROM lUserTable ORDER BY UserID DESC"
            'SQLを発行してユーザーＩＤを検索
            rsd.Open sql, conn, 3, 2, 1
            If Not rsd.EOF Then
                strInputUserID = GetNumStr(CInt(rsd("UserID"))+1, 5 )
            Else
                ' ユーザーＩＤ
                strInputUserID = "00000"
            End If

            rsd.AddNew
            rsd("UserID") = strInputUserID
            rsd("CompanyName") = strCompany
            rsd("MailAddress") = strMailAddress
            rsd.UpDate

            rsd.Close
            conn.Close

            bOK = true
            ' ユーザーＩＤをセッション変数に設定
            Session.Contents("userid") = strInputUserID
        Else
            If Trim(Request.QueryString("flg"))<>"" Then
                ' 会社名エラーのとき
                bError=true
                strError = "会社名は省略できません。"
            End If
        End If

        If Not bOK Then
            ' File System Object の生成
            Set fs=Server.CreateObject("Scripting.FileSystemobject")

            ' ログイン
            WriteLog fs, "ログイン", "新規ユーザＩＤ発行"
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
<!-------------ここからユーザーＩＤ登録画面--------------------------->
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
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <br>　
      <br>　
      <br>　
      <br>　
      <center>
      <form action="newuser.asp">
        <table>
          <tr> 
            <td nowrap>
              <dl> 
              <dt><font color="#000066" size="+1">【ユーザーＩＤ新規登録】</font><br>
              <dd>あなたのユーザーＩＤを発行しますので、以下に必要事項を入力してください。
              <dd><br>
              <dd><br>
                <table border=1 cellspacing=2 cellpadding=3 bgcolor="#FFFFFF">
                  <tr> 
                    <td nowrap bgcolor=#FFCC33><font color="#000000">会社名</font></td>
                    <td> 
                      <input type=text name=campany size=50 maxlength=200>
                     （必須入力）</td>
                  </tr>
                  <tr> 
                    <td nowrap bgcolor=#FFCC33><font color="#000000">E-mail</font></td>
                    <td> 
                      <input type=text name=mail size=30 maxlength=200>
                     （半角）</td>
                  </tr>
                </table>
              <dd><br></dl>
              <center><input type=hidden name=flg value='1'>
              <input type=submit value=" 登  録 ">
              </center>
            </td>
          </tr>
        </table>
      </form>
<%
            If bError Then
                ' エラーメッセージの表示
                DispErrorMessage strError
            End If
%>
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
<!-------------ユーザーＩＤ登録画面終わり--------------------------->
<%
    DispMenuBarBack "userchk.asp"
%>
</body>
</html>

<%
        Else
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
<!-------------ここからユーザーＩＤ登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/csvt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <br>　
      <br>　
      <br>　
      <br>　
      <center>
      <form action="
<%
                ' 戻り画面情報を取得
                strLinkID = Session.Contents("linkid")

                Response.Write strLinkID
%>
      ">
        <table>
          <tr> 
            <td nowrap>
              <dl> 
              <dt><font color="#000066" size="+1">【ユーザーＩＤ新規登録完了】</font><br>
              <dd>あなたのユーザーＩＤは[<font color="red" size="+2">
<%
                ' ユーザーＩＤの表示
                Response.Write strInputUserID
%>
                  </font>]です。
              <dd>忘れないようにメモしておいてください。
              </dl>
              <br>
              <br>
              <center>
              <input type=submit value=" 実  行 ">
              </center>
            </td>
          </tr>
        </table>
      </form>
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
<!-------------ユーザーＩＤ登録画面終わり--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
        End If
    End If
%>
