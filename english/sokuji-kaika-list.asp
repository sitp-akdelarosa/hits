<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	即時搬出システム【海貨用】	データ一覧表示
%>

<%
	' セッションのチェック
	CheckLogin "sokuji.asp"

	' 海貨コード取得
	sForwarder = Trim(Session.Contents("userid"))

	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

  ' 表示ファイルの取得
	Dim strFileName
  strFileName = Session.Contents("tempfile")

  If strFileName="" Then
		Response.Redirect "sokuji-kaika-updtchk.asp"
		Response.End
	End If
	strFileName="./temp/" & strFileName

  ' 表示ファイルのOpen
  Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

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

%>

<html>
<head>
<title>即時搬出申込み情報一覧（海貨）</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Pragma" content="no-cache">
<SCRIPT Language="JavaScript">
	function formSend(formname){
		window.document.forms[formname].submit();
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
          <td nowrap align=left> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>（海貨用）即時搬出申込み情報一覧</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>

            <br>

			<table border=0 cellpadding=0>
			  <tr>
				<td align=center colspan=2>

					<table border=0 cellpadding=0 cellspacing=2>
					  <tr>
						<td colspan=4 nowrap>
					申込みデータを更新する場合は、対象となる荷主をクリックして下さい。 <BR>
					新規に申込む場合は、'新規入力' をクリックして下さい。<BR><BR>
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
				</td>
			  </tr>
			  <tr>
				<td width=30><BR></td>
				<td nowrap>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>荷主</td>
                <td nowrap>船社</td>
                <td nowrap>船名</td>
                <td nowrap>BL／コンテナNo.</td>
                <td nowrap>対応港運</td>
                <td nowrap>対応港運<BR>TEL</td>
                <td nowrap>対応<BR>可否</td>
                <td nowrap>搬入確認予定時刻</td>
              </tr>

<%

	If LineNo>0 Then
		For i = 1 to LineNo
		    anyTmp=Split(strData(i-1),",")
%>
              <tr bgcolor="#FFFFFF"> 

				<form method=post action="sokuji-kaika-new.asp?kind=0">
					<td nowrap align=center valign=middle>
						<input type=hidden name="shipper"	 value="<%=anyTmp(9)%>">
						<input type=hidden name="shipline"	 value="<%=anyTmp(10)%>">
						<input type=hidden name="vslcode"	 value="<%=anyTmp(11)%>">

<% If Trim(anyTmp(3))<>"" Then %>
						<input type=hidden name="bl"	 value="<%=anyTmp(3)%>">
<% Else %>
						<input type=hidden name="cont"	 value="<%=anyTmp(4)%>">
<% End If %>

						<input type=hidden name="ope"		 value="<%=anyTmp(5)%>">
						<input type=hidden name="opetel"	 value="<%=anyTmp(6)%>">
						<input type=hidden name="reject"	 value="<%=anyTmp(7)%>">
						<input type=hidden name="recschtime" value="<%=anyTmp(8)%>">
						<input type=hidden name="lineno"	 value="<%=i%>">

<% If Trim(anyTmp(0))<>"" Then %>
						<a href="JavaScript:formSend(<%=i%>)"><%=anyTmp(0)%></a>
<% Else %>
						<a href="JavaScript:formSend(<%=i%>)"><%=anyTmp(9)%></a>
<% End If %>

					</td>
				</form>

<% If Trim(anyTmp(1))<>"" Then %>
				<td nowrap align=center valign=middle><%=anyTmp(1)%></td>
<% Else %>
				<td nowrap align=center valign=middle><%=anyTmp(10)%></td>
<% End If %>

<% If Trim(anyTmp(2))<>"" Then %>
				<td nowrap align=center valign=middle><%=anyTmp(2)%></td>
<% Else %>
				<td nowrap align=center valign=middle><%=anyTmp(11)%></td>
<% End If %>

<% If Trim(anyTmp(3))<>"" Then %>
				<td nowrap align=center valign=middle><%=anyTmp(3)%></td>
<% Else %>
				<td nowrap align=center valign=middle><%=anyTmp(4)%></td>
<% End If %>

<%
			For j = 0 to 8
				If anyTmp(j)=""Then
					anyTmp(j) = "<BR>"
				End If
			Next

			If Not anyTmp(8)="<BR>" Then
				anyTmp(8) = Right(anyTmp(8),11)
			End If
%>
				<td nowrap align=center valign=middle><%=anyTmp(5)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(6)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(7)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(8)%></td>
			  </tr>
<%
		Next
	End If
%>
			  <tr>
				<td nowrap align=center valign=middle>
					<a href="sokuji-kaika-new.asp?kind=1">新規入力</a>
				</td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
			  </tr>
			</table>

			<form method=get action="sokuji-kaika-updtchk.asp">
				<input type=submit value="表示データの更新">
			</form>

				</td>
			  </tr>
			</table>

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
    WriteLog fs, "7001", "即時搬出システム-海貨用情報一覧", "00", ","
%>
