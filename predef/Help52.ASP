<%@Language="VBScript" %>

<!--#include file="../Common.inc"-->

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript"><!--
function LinkSelect(form, sel)
{
        adrs = sel.options[sel.selectedIndex].value;
        if (adrs != "-" ) parent.location.href = adrs;
}
function OpenCodeWin()
{
  var CodeWin;
  CodeWin = window.open("../codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=330");
  CodeWin.focus();
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="image/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
<tr>
	<td valign=top>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td rowspan=2><img src="image/helpt52.gif" width="506" height="73"></td>
			<td height="25" bgcolor="000099" align="right"><img src="image/logo_hits_ver2.gif" width="300" height="25"></td>
		</tr>
		<tr>
			<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
' End of Addition by seiko-denki 2003.07.07
%>
			</td>
		</tr>
		</table>
		<center>
<!-- commented by seiko-denki 2003.07.07
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%> &gt; ヘルプ
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.07 -->
		<BR>
		<BR>
		<BR>
		<table>
		<tr>
			<td align="center"> 
				<table>
				<tr> 
					<td><img src="image/botan.gif" width="17" height="17" vspace="4"></td>
					<td nowrap> <b><font color="#000000">空搬出作業</font></b>&nbsp;&nbsp;</td>
					<td><img src="image/hr.gif" width="490" height="3"></td>
				</tr>
				</table>

				<table border="0" cellspacing="2" cellpadding="3">
				<tr> 
					<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 空搬出作業の新規登録</font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575">画面左側の「事前情報入力」より「空搬出」をクリックし、表示される画面に沿ってデータを登録します。<br>
					&nbsp; </td>
				</tr>
				<tr> 
					<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 空搬出作業の登録データの更新</font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575">一覧画面より修正したい作業の「ブッキング番号」をクリックし、表示される画面に沿ってデータを更新します。「ブッキング番号」をクリックした場合、表示モードにて「空搬出情報入力」画面が表示されますので、画面内の「更新モード」ボタンをクリックして、更新モードに切り替えてからデータを修正してください。<br>
					&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 検索 
					</font><font color="#666666"> </font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575">ブッキング番号による検索ができます。後方一致検索ができますので、例えば、ブッキン番号として「777]を指定し、「検索」ボタンをクリックした場合、「BOKKING000777」のブッキング番号が抽出の対象となります。<br>
					&nbsp;</td>
				</tr>
				<tr> 
				<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 「指示受諾」リンク</font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575">表示されているすべての未回答データの回答を「Yes」にします。これは、指示元がログインした場合に、指示先が指示内容を受諾したかどうかの判断に利用します。指示元に対して「No」と回答する場合は、「作業番号」をクリックし、「空搬出情報入力（表示モード）」画面から「空搬出情報入力（更新モード）」画面へ切り替えて、「指示元へ回答」から「No」を選択して、「更新」ボタンをクリックします。<br>
					&nbsp;</td>
				</tr>
				<tr> 
				<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 「ＣＳＶファイル出力」リンク</font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575">表示されている一覧のデータをＣＳＶファイルで取り出すことができます。<br>
					&nbsp;</td>
				</tr>
				<tr> 
				<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 空搬出作業の一覧表示</font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575">画面左側の「表示種類切替」より「空搬出」をクリックすると空搬出作業一覧が表示されます。<br>
					&nbsp;</td>
				</tr>
				<tr> 
				<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 一覧の表示項目、表示順の切替</font></b></td>
				</tr>
				<tr> 
					<td width="15"> </td>
					<td width="575"><form>
						<dt> 以下のようなセレクトボックスで一覧の表示項目、表示順の切替を行います。<br>
						<dd>
						<table border="0" cellspacing="0" cellpadding="0">
						<tr> 
							<td valign="middle" nowrap><font color="#FF0033">【例】</font></td>
							<td> 
								<table border="1" cellspacing="1" cellpadding="5">
									<tr> 
									<td bgcolor="#FFFFFF" align="center" valign="middle">
										<select>
										<option>入力日順に表示</option>
										<option>指示先が未回答のブッキング一覧</option>
										<option>指示先回答がＮｏのブッキング一覧</option>
										<option>搬出未完了分をすべて表示</option>
										<option>全件表示</option>
										</select>
									</td>
								</tr>
								</table>
							</td>
						</tr>
						</table>
						<br>
						<table border="1" cellspacing="0" cellpadding="2" bgcolor="#FFFFFF">
						<tr bgcolor="#99aaFF" align="center"> 
							<td nowrap><b><font color="#333333">選択項目</font></b></td>
							<td width="300"><b><font color="#333333">説明</font></b></td>
						</tr>
						<tr valign="top"> 
						<td nowrap>入力日順に表示</td>
<%'						<td>入力日の昇順に一覧を表示します。</td>%>
						<td>入力日の降順に一覧を表示します。</td>
						</tr>
						<tr valign="top"> 
						<td nowrap>指示先が未回答の<br>ブッキング一覧</td>
						<td>作業の指示先が照会していない（データの更新をしていない、「指示受諾」ボタンのクリックをしていない、または「No」と回答していない）ブッキングの一覧を表示します。</td>
						</tr>
						<tr valign="top"> 
						<td nowrap>指示先回答がＮｏの<br>ブッキング一覧</td>
						<td>作業の指示先が「No」と回答しているブッキングの一覧を表示します。</td>
						</tr>
						<tr valign="top"> 
						<td nowrap>搬出未完了分を<br>すべて表示</td>
<%'						<td>完了日時が設定されていない作業をすべて表示します。</td>%>
						<td>ピック済み本数がピック予約本数に達していない作業をすべて表示します。</td>
						</tr>
						<tr valign="top"> 
						<td nowrap>全件表示</td>
						<td>作業完了分も含めて、登録されている作業をすべて表示します。</td>
						</tr>
						</table>
						</form>
 					
					
					<br>
					&nbsp;</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
<!---------->
		</center>
	</td>
</tr>
<!-- Commented 2003.9.8
<tr>
	<td valign="bottom"> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td valign="bottom" align="right"><a href="http://www.hits-h.com/index.asp"><img src="image/b-home.gif" border="0" width="105" height="23" usemap="#map"></a></td>
		</tr>
		<tr>
			<td bgcolor="000099" height="10"><img src="image/1.gif"></td>
		</tr>
		</table>
	</td>
</tr>
Commented 2003.9.8 -->
</table>
<!-------------画面終わり--------------------------->
<!-- Commented 2003.9.8
<map name="map"> 
  <AREA coords=22,0,0,22,105,22,105,0 href="http://www.hits-h.com/index.asp" shape=POLY>
</map>
Commented 2003.9.8 -->
</body>
</html>