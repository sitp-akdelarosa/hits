<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'
	'	【コンテナ情報入力】	入力画面
	'
%>

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"
	' 海貨コード
	sSosin = Trim(Session.Contents("userid"))

	' 新規(1) or 更新(0)
    bKind = Request.QueryString("kind")
	Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sCont,sSeal,sCargoWeight,sContWeight,sRifer,sDanger
	sUser 		= Request.form("user")
	sUserNo 	= Request.form("userno")
	sVslCode 	= Request.form("vslcode")
	sVoyCtrl 	= Request.form("voyctrl")
	sBooking 	= Request.form("booking")
	If bKind=0 Then
		sCont 		= Request.form("cont")
		sSeal 		= Request.form("seal")
		sCargoWeight= Request.form("cargow")
		sContWeight	= Request.form("contw")
		sRifer 		= Request.form("rifer")
		sDanger 	= Request.form("danger")
	End If
	iLineNo		= Request.form("lineno")

%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

	function checkFormValue(){
		contvalue = window.document.input.cont.value;
		if(contvalue == ""){
			window.alert("コンテナNo.が未入力です。");
			return false;
		}
		return true;
	}

	// 数値チェック
	function checknum(etext)
	{
		if (etext.value == "")
			return false;

		if (isNaN(etext.value)) {
			alert("数値を入力して下さい。");
			etext.focus();
			etext.select();
			return false;
		}

		fTemp=parseFloat(etext.value)
	    if (fTemp>99.9) {
			alert("99.9Ton以下の数値を入力して下さい。");
			etext.focus();
			etext.select();
			return false;
		}

		return true;
	}

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
          <td rowspan=2><img src="gif/kaika5t.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>更新情報入力</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>
      <table>
        <tr>
          <td nowrap align=center>
				輸出コンテナについて、以下の項目を入力して送信を押して下さい。
            <form method=post name="input" action="ms-kaika-expcontinfo-exec.asp">
				<input type=hidden name="kind" value="<%=bKind%>">
				<input type=hidden name="lineno" value="<%=iLineNo%>">
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>船名</b></font>
                  </td>
                  <td nowrap>
                    <%=sVslCode%>
					<input type=hidden name="vslcode" value="<%=sVslCode%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Voyage No.</b></font>
                  </td>
                  <td nowrap>
                    <%=sVoyCtrl%>
					<input type=hidden name="voyctrl" value="<%=sVoyCtrl%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
					 <font color="#FFFFFF"><b>荷主コード</b></font>
				  </td>
                  <td nowrap>
                    <%=sUser%>
					<input type=hidden name="user" value="<%=sUser%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>荷主管理番号</b></font>
                  </td>
                  <td nowrap>
                    <%=sUserNo%>
					<input type=hidden name="userno" value="<%=sUserNo%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Booking No.</b></font>
                  </td>
                  <td nowrap>
                    <%=sBooking%>
					<input type=hidden name="booking" value="<%=sBooking%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>コンテナNo.</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=cont value="<%=sCont%>" size=14 maxlength=12>
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
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>シールNo.</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=seal value="<%=sSeal%>" size=17 maxlength=15>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>貨物重量</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=cargow value="<%=sCargoWeight%>" size=5 maxlength=4 onblur="checknum(document.input.cargow)">（t）
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ 半角数値 ]</font>
						</td>
					  </tr>
					</table>
                    
					&nbsp;&nbsp;&nbsp;<font size="-1">小数点以下1桁まで有効&nbsp;&nbsp;（例）10.2</font>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>総重量</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=contw value="<%=sContWeight%>" size=5 maxlength=4 onblur="checknum(document.input.contw)">（t）
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ 半角数値 ]</font>
						</td>
					  </tr>
					</table>
                    
					&nbsp;&nbsp;&nbsp;<font size="-1">小数点以下1桁まで有効&nbsp;&nbsp;（例）10.2</font>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>リーファー</b></font>
                  </td>
                  <td nowrap>
<%	
	Dim strRifKind
	If bKind=0 And sRifer="R" Then
		strRifKind = "checked"
	End If
%>
					<input type=checkbox name=rifer <%=strRifKind%>>
					<font size=-1>リーファーの場合チェックして下さい。</font>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>危険物</b></font>
                  </td>
                  <td nowrap>
<%	
	Dim strDngKind
	If bKind=0 And sDanger="H" Then
		strDngKind = "checked"
	End If
%>
					<input type=checkbox name=danger <%=strDngKind%>>
					<font size=-1>危険物の場合チェックして下さい。<sup>（※）</sup></font>
                  </td>
                </tr>

              </table>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font size=-1>（※） 消防法に関わる危険物の場合のみチェックして下さい。</font>
              <br><BR>
                <input type=submit name="send" value=" 送  信 " onClick="return checkFormValue()">
                <input type=button value=" 中  止 " onClick="JavaScript:window.history.back()">

				</center>
              </center>
            </form>
<%
            If bError Then
                ' エラーメッセージの表示
                DispErrorMessage strError
            End If
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
    DispMenuBarBack "ms-kaika-expcontinfo.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    WriteLog fs, "4106","海貨入力輸出コンテナ情報-情報入力","00", ","
%>
