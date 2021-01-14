<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	即時搬出システム【港運用】	データ一覧表示

%>

<%
	' セッションのチェック
	CheckLogin "sokuji.asp"

	' 港運コード取得
	sOpe = Trim(Session.Contents("userid"))

	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' テンポラリファイル名を作成して、セッション変数に設定
	Dim strFileName
	strFileName = Session.Contents("tempfile")

  If strFileName="" Then
		Response.Redirect "sokuji-koun-updtchk.asp"
		Response.End
	End If
	strFileName="./temp/" & strFileName

  ' 表示ファイルのOpen
  Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' テンポラリファイルの読み込み
	Dim strData()
	LineNo=0
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)
	Do While Not ti.AtEndOfStream
		  strTemp=ti.ReadLine
		  ReDim Preserve strData(LineNo)
		  strData(LineNo) = strTemp
		  LineNo=LineNo+1
	Loop
	ti.Close
	Session.Contents("ChkCount")=LineNo

%>
<html>
<head>
<title>即時搬出申込み情報一覧（港運）</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Pragma" content="no-cache">
<SCRIPT Language="JavaScript">
	function checkFormValue(){
<%
	If LineNo>0 Then
		For i=1 to LineNo
			If LineNo=1 Then
				Response.Write "if(document.koun.chk" & i & ".checked==false)"
			ElseIf i=1 Then
				Response.Write "if((document.koun.chk" & i & ".checked==false)"
			ElseIf i=LineNo Then
				Response.Write "&&(document.koun.chk" & i & ".checked==false))"
			Else
				Response.Write "&&(document.koun.chk" & i & ".checked==false)"
			End If
		Next
%>
		{ return showAlert("チェック",true); }
		return true;
<%
	End If
%>
	}
	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "が未入力です。");
		} else {
			window.alert(strAlert + "が不正です。");
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
          <td rowspan=2><img src="gif/sokuji2t.gif" width="506" height="73"></td>
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

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
<%
	strNowTime = Year(Now) & "年" & _
		Right("0" & Month(Now), 2) & "月" & _
		Right("0" & Day(Now), 2) & "日" & _
		Right("0" & Hour(Now), 2) & "時" & _
		Right("0" & Minute(Now), 2) & "分現在の情報"

%>
					&nbsp;&nbsp;<%=strNowTime%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>（港運用）即時搬出申込み情報一覧</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>
			<center>

            <br>
			<table border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td nowrap align=left >
			対応可能な場合は目的のデータ（右の四角の枠内）にチェックして予定時刻入力を押して下さい。<BR>
			対応不可の場合は目的のデータ（右の四角の枠内）にチェックして対応不可を押して下さい。<BR><BR>
				</td>
			  </tr>
			</table>

					<table border=0 cellpadding=0 cellspacing=2 width=500>
					  <tr>
						<td colspan=4 nowrap>
							＜実証実験の実施方法＞
						</td>
					  </tr>
					  <tr>
						<td width=20 rowspan=5><BR></td>
						<td nowrap valign=top>
							１．海貨 → ターミナルの連絡
						</td>
						<td valign=top nowrap> ： </td>
						<td valign=top>
							事前に対象の船名、Voyage No.、コンテナNo.をターミナルに電話で連絡する
						</td>
					  </tr>
					  <tr>
						<td nowrap valign=top>
							２．海貨 → ターミナルの申し込み
						</td>
						<td valign=top nowrap> ： </td>
						<td valign=top>
							Web画面上に入力と同時に電話で連絡
						</td>
					  </tr>
					  <tr>
						<td nowrap valign=top>
							３．ターミナル → 海貨の回答
						</td>
						<td valign=top nowrap> ： </td>
						<td valign=top>
							Web画面上に入力と同時に電話で連絡
						</td>
					  </tr>
					  <tr>
						<td nowrap colspan=3>
							４．OKなら、海貨経由で担当する陸運会社がHITSでシャトル便を予約<BR>
							５．シャトル便でコンテナ搬出
						</td>
					  </tr>
					</table>
					<BR>

			<table border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td align=center nowrap>

					<table border=0 cellpadding=0 cellspacing=0>
					<tr>
					<td nowrap align=right>

				    <form method=post action="sokuji-koun-new.asp" name="koun">

		            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
		              <tr align="center" bgcolor="#FFCC33"> 
		                <td nowrap>海貨</td>
		                <td nowrap>船社</td>
		                <td nowrap>船名</td>
		                <td nowrap>BL／コンテナNo.</td>
		                <td nowrap>対応<BR>可否</td>
		                <td nowrap>搬入確認<BR>予定時刻</td>
		                <td nowrap><BR></td>
		              </tr>

<%
	For i = 1 to LineNo
	    anyTmp=Split(strData(i-1),",")

		For j = 5 to 6
			If anyTmp(j)="" Then
				anyTmp(j) = "<BR>"
			End If
		Next

		If Not anyTmp(6)="<BR>" Then
			anyTmp(6) = Right(anyTmp(6),11)
		End If
%>
		              <tr bgcolor="#FFFFFF"> 

<% If Trim(anyTmp(0))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(0)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(7)%></td>
<% End If %>
<% If Trim(anyTmp(1))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(1)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(8)%></td>
<% End If %>
<% If Trim(anyTmp(2))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(2)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(9)%></td>
<% End If %>
<% If Trim(anyTmp(3))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(3)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(4)%></td>
<% End If %>
						<td nowrap align=center valign=middle><%=anyTmp(5)%></td>
						<td nowrap align=center valign=middle><%=anyTmp(6)%></td>
						<td nowrap align=center valign=middle>
						  <input type=checkbox name=chk<%=i%>>
						</td>

<%
		If anyTmp(6)="<BR>" Then anyTmp(6)=""
		If anyTmp(7)="<BR>" Then anyTmp(7)=""
		If anyTmp(8)="<BR>" Then anyTmp(8)=""
		If anyTmp(8)="<BR>" Then anyTmp(9)=""
		If anyTmp(3)="<BR>" Then anyTmp(3)=""
		If anyTmp(4)="<BR>" Then anyTmp(4)=""
		If anyTmp(5)="<BR>" Then anyTmp(5)=""
%>
						<input type=hidden name=shipper<%=i%> value=<%=anyTmp(7)%>>
						<input type=hidden name=shipline<%=i%> value=<%=anyTmp(8)%>>
						<input type=hidden name=vslcode<%=i%> value=<%=anyTmp(9)%>>
						<input type=hidden name=forwarder<%=i%> value=<%=anyTmp(10)%>>
<% If Trim(anyTmp(3))<>"" Then %>
						<input type=hidden name=bl<%=i%> value=<%=anyTmp(3)%>>
<% Else %>
						<input type=hidden name=cont<%=i%> value=<%=anyTmp(4)%>>
<% End If %>
						<input type=hidden name=reject<%=i%> value=<%=anyTmp(5)%>>
						<input type=hidden name=recschtime<%=i%> value=<%=anyTmp(6)%>>

					  </tr>
<%
	Next
%>
					</table>
					<BR>
					<div align=left>
					<input type=button value="表示データの更新" onclick="window.location.href='sokuji-koun-updtchk.asp'">
					</div>
					<input type=submit name=timeset value="予定時刻入力" onClick="return checkFormValue()">
					<input type=submit name=corrfail value=" 対 応 不 可 " onClick="return checkFormValue()">

					</td>
					</tr>

					</form>

					<tr><td align="left">
					</td></tr>

					</table>

				</td>
			  </tr>
			</table>

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
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    WriteLog fs, "7003", "即時搬出システム-港運用情報一覧", "00", ","
%>
