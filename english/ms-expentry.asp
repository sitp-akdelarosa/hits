<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' DBの接続
    ConnectSvr conn, rsd

    ' ユーザ種類を取得する
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' セッションが切れているとき
        Response.Redirect "expentry.asp"             '輸出コンテナ照会トップ
        Response.End
    End If

	Dim iNum
	If strUserKind="海貨" Then
		iNum = 1101
	ElseIf strUserKind="陸運" Then
		iNum = 1102
	Else
		iNum = 1103
	End If
    ' 輸出入業務支援-輸出コンテナ照会
    WriteLog fs, iNum,"輸出コンテナ照会-" & strUserKind & "用照会","00", ","
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
	sVslCode=document.con.vessel.value;
	sVoyCode=document.con.voyage.value;
	if ((sVslCode!="" && sVoyCode=="")||(sVslCode=="" && sVoyCode!="")) {	/* 船のチェック */
			window.alert("船名(コールサイン)とVoyage No.はペアで入力してください。");
			return false;
	}
	return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから照会画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<td rowspan=2><%

    If strUserKind="海貨" Then
        Response.Write "<img src='gif/expkaika.gif' width='506' height='73'>"
    ElseIf strUserKind="陸運" Then
        Response.Write "<img src='gif/exprikuun.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/expninushi.gif' width='506' height='73'>"
    End If

%></td>
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
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>輸出コンテナ情報照会
<%
    If strUserKind="海貨" Then
        Response.Write "(海貨用)"
    ElseIf strUserKind="陸運" Then
        Response.Write "(陸運用)"
    Else
        Response.Write "(荷主用)"
    End If
%>
            </b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
        <tr>
        </tr>

      </table>
      <table width="480">
        <tr>
          <td colspan="4">
			条件を入力しないで輸出照会ボタンを押すと、全てのデータが表示されます。<BR><BR>
            データが多い場合は表示できない事がありますので、
			その場合は下記フォームに適当な照会条件を入力し、
			輸出照会ボタンを押して下さい。
          </td>
        </tr>
      </table>
<%
    If strUserKind<>"陸運" Then
%>
      <form name="con" method="get" action="ms-expcntnr.asp" onSubmit="return ClickSend()">
<%
    Else
%>
      <form name="con" method="get" action="ms-expcntnr.asp">
<%
    End If
%>
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
<%
    If strUserKind<>"陸運" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap>
                    <table border=0 cellpaddig=0 cellspacing=0>
                      <tr><td><font color="#FFFFFF"><b>船名(コールサイン)</b></font></td></tr>
                      <tr><td><font color="#FFFFFF"><b>Voyage No.</b></font></td></tr>
                    </table>
                    </td>
                  <td nowrap>
                    <table border=0 cellpaddig=0 cellspacing=0>
                    <tr>
						<td width=150><input type=text name=vessel size=10 maxlength="7"></td>
						<td><font size="1" color="#2288ff">[半角英数]</font></td>
					</tr>
                    <tr>
						<td width=150><input type=text name=voyage size=18 maxlength="12"></td>
						<td><font size=1 color="#2288ff">[半角英数]</font></td>
					</tr>
                    </table>
                  </td>
                </tr>
<%
    End If
    If strUserKind="海貨" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>荷主コード</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
		                    <input type=text name=ninushi size=8 maxlength="5"> 
						</td>
						<td>
							<font size=1 color="#2288ff">[半角英数]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>指定陸運業者コード</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
		                    <input type=text name=rikuun size=5 maxlength="3">
						</td>
						<td>
							<font size=1 color="#2288ff">[半角英数]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<%
    End If
%>
<%
    If strUserKind<>"海貨" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>海貨コード</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=kaika size=8 maxlength="5">
						</td>
						<td align=right valign=middle nowrap>
							<font size=1 color="#2288ff">[半角英数]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<%
    End If
%>
              </table>
              <br>
              <input type=submit value="   輸出照会   ">
      </form>
      <br>
      <br>
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
<!-------------照会画面終わり--------------------------->
<%
    DispMenuBarBack "expentry.asp"
%>
</body>
</html>
