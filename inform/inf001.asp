<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf001.asp                                      _/
'_/	Function	:お知らせ一覧表示                                _/
'_/	Date			:2005/03/07                                      _/
'_/	Code By		:aspLand HARA                                    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Hakata Port IT Systemお知らせ</title>
<link href="style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT language="JavaScript">
<!--
function openwin(){
	f=document.information;
	Win = window.open('inf100.asp','email_regist','left=100,top=100,width=500,height=150,resizable=yes,scrollbars=no,status=yes');
}
-->
</SCRIPT>
</head>
<!--#include File="common.inc"-->
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="images/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<table width="50%" border="0">
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td align="center"><font class="header3">お知らせ一覧</font></td>
	</tr>
	<tr><td height="20">&nbsp;</td></tr>
</table>
<form name="information">
<table width="50%" border="0">
	<tr>
		<td colspan="2">
			●表示のためには「AcrobatReader」というソフトをインストールする必要があります。
		</td>
	</tr>
	<tr>
		<td width="30">&nbsp;</td>
		<td>
			<table>
				<tr>
					<td valign="top">ダウンロードはこちらから－－－－→</td>
					<td>
						<map name="nE8A.answer.0.2B44">
						<area shape=rect coords="-2,0,87,29" is="HotspotRectangle20_1" href="http://www.adobe.co.jp/products/acrobat/readstep2.html" alt="Get Adobe Reader" target="_blank"></map>
						<img src="images/AcrobatReader.gif" width="88" height="31" usemap="#nE8A.answer.0.2B44" border="0">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			●新着の連絡mailをご希望の方はmailアドレスをご登録ください。<br>
		</td>
	</tr>
	<tr>
		<td width="30">&nbsp;</td>
		<td>
			mailアドレスの登録・更新・削除は<a href="javascript:openwin();">こちら</a>
		</td>
	</tr>
</table>
<br>
<%
	Dim param(2), fso, fod, fic, home_path, rel_path
	Dim cnt, file_info, i, j, k, work

	'''iniファイルの値の読み込み
	getIni param

	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set fod=fso.GetFolder(param(0))
	set fic=fod.Files
	cnt=0
	For Each fil In fic
		cnt=cnt+1
	Next

	ReDim file_info(cnt)
	'''ホームディレクトリの絶対物理パスの取り出し
	home_path=Request.ServerVariables("APPL_PHYSICAL_PATH")
	'''ホームディレクトリの相対パス
	rel_path="/" & Replace(Right(param(0),len(param(0))-len(home_path)),"\","/")

	i=0
	'''file_info配列にファイルの作成日と名前を格納
	For Each fil In fic
		if(DateDiff("d",fil.DateLastModified,Date)<=CInt(param(1))) then '''今日－作成日<=特定期間
			file_info(i)=Left(fil.DateLastModified,4) & "年" & Mid(fil.DateLastModified,6,2) & "月" & Mid(fil.DateLastModified,9,2) & "日" & ":" & fil.Name & ":1"
		else		'''今日－作成日>特定期間
			file_info(i)=Left(fil.DateLastModified,4) & "年" & Mid(fil.DateLastModified,6,2) & "月" & Mid(fil.DateLastModified,9,2) & "日" & ":" & fil.Name & ":0"
		end if
		i=i+1
	Next
	file_num=i
	f=Array(0,0)
	ReDim f(file_num,3)
	'''作成日の新しいものがより上に表示されるようにソートする
	For i = 0 To UBound(file_info) - 1
		For j = i + 1 To UBound(file_info)
			If StrComp(file_info(i),file_info(j),1)<0 Then '''file_info(i)がfile_info(j)より小さい
				work=file_info(i)
				file_info(i)=file_info(j)
				file_info(j)=work
			End If
		Next
	Next

	IF file_num >0 then
%>
		<table width="600" border="0" cellspacing="0" cellpadding="0">
			<tr valign="top">
				<td width="1" bgcolor="red"><img src="images/spacer_FF0000.gif" width="1" height="1" border="0"></td>
				<td width="598">
					<table width="598" border="0" cellspacing="0" cellpadding="0" bgcolor="red">
						<tr>
							<td height="1"><img src="images/spacer_FF0000.gif" width="1" height="1" border="0"></td>
						</tr>
					</table>
					<table width="498" border="0" cellspacing="2" cellpadding="0">
						<tr valign="top">
							<td width="20%" align="center">登録日</td>
							<td width="80%" align="left">　　　　内容</td>
						</tr>
<%
		'''作成日とファイル名の切り出し
		For k=0 to file_num-1
			file_data=split(file_info(k),":")
			j=0
			for each fd in file_data
				f(k,j)=fd
				j=j+1
			next
%>
						<tr>
							<td colspan="2">
								<a href="<%=rel_path&f(k,1)%>" target="_blank"><%= f(k,0) %>　<%= left(f(k,1),len(f(k,1))-4) %></a>
<%
			if f(k,2)=1 then
%>
					　<img src="./images/new2.gif" border="0">
<%
			end if
%>
							</td>
						</tr>
<%
		Next
%>
					</table>

					<table width="598" border="0" cellspacing="0" cellpadding="0" height="1">
						<tr>
							<td bgcolor="red"><img src="images/spacer_FF0000.gif" width="1" height="1" border="0"></td>
						</tr>
					</table>
				</td>
				<td width="1" bgcolor="red"><img src="images/spacer_FF0000.gif" width="1" height="1" border="0"></td>
			</tr>
		</table>

<%
	ELSE
		response.write("<table width='50%' border='0'>")
		response.write("<tr><td>&nbsp;</td></tr>")
		response.write("<tr><td align='center'><font color='red'>お知らせはありません。</font></td></tr>")
		response.write("</table>")
	END IF

%>

<br>
<br>

<font class="font10"><a href="../index.asp">TOPに戻る</a></font>
</center>
</form>
</body>
</html>
