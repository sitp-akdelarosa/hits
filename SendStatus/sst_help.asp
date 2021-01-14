<%@Language="VBScript" %>

<!--#include file="../Common.inc"-->

<html>
<head>
<title>ステータス配信依頼ヘルプ</title>
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
<table border="0" cellspacing="0" cellpadding="0" width="100%" height=100%>
<tr>
	<td valign=top>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td rowspan=2><img src="image/sst_help.gif" width="506" height="73"></td>
			<td height="25" bgcolor="000099" align="right"><img src="image/logo_hits_ver2.gif" width="300" height="25"></td>
		</tr>
		<tr>
			<td align="right" width="100%" height="48"> 
<%
call	DisplayCodeListButton
%>
			</td>
		</tr>
		</table>
		<center>
		<BR><BR><BR>
		<table border="0">
			<tr>
				<td align="center"> 
					<table border="0" cellspacing="2" cellpadding="3">
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ ステータス配信の新規登録</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">画面左側の「表示種類切替」より「新規依頼」をクリックし、表示される画面にてコンテナ番号またはＢＬ番号を入力し、
															「登録」ボタンをクリックします。<br>
															コンテナヤード搬出済のコンテナ番号も登録できますが、搬出後１１日以上経過したものは登録できません。
															また、ＢＬ番号指定の場合、関連するコンテナ番号がすべて搬出後１１日以上経過したものは登録できません。<br>
															ＢＬ番号指定の場合、ＨｉＴＳ登録後にそのＢＬに追加されたコンテナについてはステータスを送信できません。
															（対策）ステータス配信からそのＢＬの登録を一旦削除し、再度登録します。</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr>
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ ステータス配信依頼中一覧表示</font></b></td>
						</tr>
						<tr>
							<td width="15">　</td>
							<td width="575">画面左側の「表示種類切替」より「依頼中一覧」をクリックするとステータス配信依頼中一覧が表示されます。
															なお、コンテナヤード搬出後１１日以上経過したコンテナ番号は一覧には表示されません。また、ＢＬ番号指定の場合、
															すべてのコンテナが搬出後１１日以上経過したものは一覧には表示されません。</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ ステータス配信の登録データの削除</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">ステータス配信依頼中一覧画面より削除したい「No.」をクリックし、
															表示される画面にて「削除」ボタンをクリックします。</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ mail即時送信</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">新規登録画面にて、コンテナ番号またはＢＬ番号を入力して「mail即時送信」をクリックすると、
															指定したコンテナまたはＢＬに関連するコンテナの現在の状態がメールにて送信されます。<br>
															また、一覧から項番のクリックにより表示される詳細画面にて、「mail即時送信」をクリックすると、
															同様にコンテナまたはＢＬに関連するコンテナの現在の状態がメールにて送信されます。<br>
															なお、メールにて送信するには輸入ステータス配信依頼（設定）画面にてメールアドレスを登録しておく必要があります。
															また、mail即時送信では、輸入ステータス配信依頼（設定）画面の設定内容に係わりなく、すべての項目についてメール送信されます。</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ ステータス配信対象項目の設定</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">画面左側の「表示種類切替」より「設定」をクリックすると輸入ステータス配信依頼（設定）画面が表示されます。
															状態が変化した場合にメールが送信されてくるように設定したい項目を選んで、「登録」ボタンをクリックします。
															当該画面にて登録されたメールアドレスへメールが送信されます。</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ 検索</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">コンテナ番号またはＢＬ番号による検索ができます。また、後方一致検索ができます。
															例えば、コンテナ番号として「555]を指定し、「検索」ボタンをクリックした場合、
															「CONT0000555」のコンテナ番号は抽出の対象となります。<br>
															コンテナ番号検索後、元に戻すときは、左側の「依頼中一覧」をクリックしてください。</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<table border="0">
			<form>
			<tr><td>　</td></tr>
			<tr><input type="button" value="閉じる" onClick="window.close()"></td></tr>
			</form>
		</table>
		</center>
	</td>
</tr>
</table>
<!-------------画面終わり--------------------------->
</body>
</html>
