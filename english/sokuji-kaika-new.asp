<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	即時搬出システム【海貨用】	変更,削除用画面
%>

<%
	' セッションのチェック
	CheckLogin "sokuji.asp"

	' 海貨コード
	sForwarder = Trim(Session.Contents("userid"))

	' 新規追加後(2) or 新規(1) or 更新(0)
	Dim bKind
	bKind = Request.QueryString("kind")

	If bKind=0 Then
		Session.Contents("kind") = 0
	ElseIf bKind=1 Then
		Session.Contents("kind") = 1
	ElseIf bKind=2 Then
		Session.Contents("kind") = 2
	End If

	If bKind = 0 Then
		Dim sShipper,sShipLine,sVslCode,sBL,sCont,sReject,sRecschTime,iLineNo
		sShipper 	= Request.form("shipper")
		sShipLine 	= Request.form("shipline")
		sVslCode 	= Request.form("vslcode")
		sOpe 		= Request.form("ope")
		sOpeTel		= Request.form("opetel")
		sBL 		= Request.form("bl")
		sCont 		= Request.form("cont")
		sReject 	= Request.form("reject")
		sRecschTime = Request.form("recschtime")
		iLineNo		= Request.form("lineno")
	End If

%>

<html>
<head>
<title>即時搬出申込み（海貨）</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function checkFormValue(){
		if(!checkBlank(getFormValue(0))){ return showAlert("荷主コード",true); }
		if(!checkBlank(getFormValue(1))){ return showAlert("船社コード",true); }
		if(!checkBlank(getFormValue(2))){ return showAlert("船名コード",true); }
		if(!checkBlank(getFormValue(3)) && !checkBlank(getFormValue(4))){ return showAlert("BL No.またはコンテナNo.",true); }
		if(checkBlank(getFormValue(3)) && checkBlank(getFormValue(4))){ return showAlert("BL No.とコンテナNo.",false); }
		return true;
	}
	function getFormValue(iNum){
		formvalue = window.document.input.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}
	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "が未入力です。");
		} else {
			window.alert(strAlert + "は、どちらか一方を入力して下さい。");
		}
		return false;
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
          <td rowspan=2><img src="gif/sokuji1t.gif" width="506" height="73"></td>
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
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>（海貨用）即時搬出申込み</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>

              <center>
            <br>
			即時搬出対象貨物について、下の各項目を入力して下さい。

            <form method=post name="input" action="sokuji-kaika-exec.asp">
			<table border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td nowrap align=left>


              <table border="1" cellspacing="2" cellpadding="2" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" width=120 align=center valign=middle>
                    <font color="#FFFFFF"><b>荷主コード</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180 height=25 valign=middle>
<% If bKind=1 Then %>
							<input type=text name=shipper value="<%=sShipper%>" size=7 maxlength=5>
<% Else %>
							<font>&nbsp;<%=sShipper%></font>
							<input type=hidden name="shipper" value="<%=sShipper%>">
<% End If %>
						</td>
						<td align=left valign=middle nowrap>
<% If bKind=1 Then %>
							<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
<% End If %>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>船社コード</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180 height=25 valign=middle>
<% If bKind=1 Then %>
							<input type=text name=shipline value="<%=sShipLine%>" size=7 maxlength=5>
<% Else %>
							<font>&nbsp;<%=sShipLine%></font>
							<input type=hidden name="shipline" value="<%=sShipLine%>">
<% End If %>
						</td>
						<td align=left valign=middle nowrap>
<% If bKind=1 Then %>
							<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
<% End If %>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>船名コード</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180 height=25 valign=middle>
<% If bKind=1 Then %>
							<input type=text name=vslcode value="<%=sVslCode%>" size=9 maxlength=7>
<% Else %>
							<font>&nbsp;<%=sVslCode%></font>
							<input type=hidden name="vslcode" value="<%=sVslCode%>">
<% End If %>
						</td>
						<td align=left valign=middle nowrap>
<% If bKind=1 Then %>
							<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
<% End If %>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

			  </table>

			  <BR><BR>

				<center>BL No.または、コンテナNo.のどちらかを入力して送信ボタンを押して下さい。</center>
				<BR>

              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" width=120 align=center valign=middle>
                    <font color="#FFFFFF"><b>BL No.の場合</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180>
							<input type=text name=bl value="<%=sBL%>" size=22 maxlength=20>
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
                    <font color="#FFFFFF"><b>コンテナNo.の場合</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180>
							<input type=text name=cont value="<%=sCont%>" size=14 maxlength=12>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

              </table>
              <br>
				<center>

<% If bKind=0 Then %>
				<input type=hidden name="blold" value="<%=sBL%>">
				<input type=hidden name="contold" value="<%=sCont%>">
<% End If %>

				<input type=hidden name="ope" value="<%=sOpe%>">
			  <input type=hidden name="opetel" value="<%=sOpeTel%>">
			  <input type=hidden name="reject" value="<%=sReject%>">
			  <input type=hidden name="recschTime" value="<%=sRecschTime%>">
			  <input type=hidden name="lineno" value="<%=iLineNo%>">
              <input type=submit name="send" value=" 送  信 " onClick="return checkFormValue()">
              <input type=submit name="stop" value=" 終  了 ">

<% If bKind<>1 Then %>

              <input type=submit name="del" value=" 削  除 ">

<% End If %>

				</center>
				</td>
			  </tr>
			</table>
            </form>
              </center>
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
    DispMenuBarBack "sokuji-kaika-list.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    If bKind=0 Then
    	WriteLog fs,"7002", "即時搬出システム-海貨用申込み", "02", ","
	Else
	    WriteLog fs,"7002", "即時搬出システム-海貨用申込み", "01", ","
	End If
%>
