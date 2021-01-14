<%@ LANGUAGE="VBScript" %>
<!--#include file="Common.inc"-->
<html>

<head>
<title>運行ダイヤ一覧画面</title>
</head>

<body>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title02.gif" width="236" height="34"><p>
<b><font size=5>○○○○ユーザー向け画面　　○○○○グループ</font></b><br>
<b><font size=5>運行ダイヤ一覧</font></b>
<br>
</center>
<br>   
   

		<span style="background-color:white;">
		<p>
		<form>
		<center>
		<table rules="none" width="200"  >   
					
			<tr>
			    <td align=center>運行ダイヤ照会</td>			
			</tr>

			<tr>
				<td align=center><INPUT TYPE="radio" NAME="radkaitou" VALUE="1">当日
								 <INPUT TYPE="radio" NAME="radkaitou" VALUE="2">翌日</td>
			</tr>

			<tr>
				<td align=center><input type="submit" value="照会" id=submit4 name=submit4></td>
			</tr>


		</table>
		</center>
		</form>
		</p>

		<span style="background-color:white;">

		
		<br>
		<center><b>2000年12月13日運行ダイヤ　9:50現在</b><br>
			不足空シャーシ（20'･･･３本　40'･･･２本）　　空スロット（３本）
		</center>

		<font face="ＭＳ ゴシック">
		<center>
		<table border="1" width="570"  >   

			<tr>
			    <td bgcolor="#e8ffe8">時間帯</td>
			    <td bgcolor="#e8ffe8">(作業番号)　コンテナ番号　　種類　　ｻｲｽﾞ　　状態　　　　場所</td>
			</tr>

			<tr>
			    <td rowspan=5>9:00</td>
			    <td><a href="SYWB016.asp">(13001)　　　MESU00000001　　入　　　20　　　完了</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13002)　　　MESU00000002　　出　　　20　　　完了</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13005)　　　MESU00000003　　入　　　40　　　完了　　　　ＳＹ</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13007)　　　MESU00000004　　出　　　40　　　完了　　　　ＳＹ</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13009)　　　MESU00000005　　入　　　20　　　完了　　　　ＣＹ</a></td>
			</tr>

			<tr>
			    <td rowspan=3>10:00</td>
			    <td><a href="SYWB016.asp">(13010)　　　MESU00000006　　入　　　40　　　完了</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13011)　　　MESU00000007　　出　　　40　　　完了</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13012)　　　MESU00000008　　入　　　20　　　完了　　　　ＳＹ</a></td>
			</tr>

			<tr>
			    <td rowspan=2>11:00</td>
			    <td><a href="SYWB016.asp">(13015)　　　MESU00000011　　入　　　40　　　完了　　　　ＳＹ</a></td>
			</tr>

			<tr>
			    <td><a href="SYWB016.asp">(13016)　　　MESU00000013　　出　　　20　　　完了</a></td>
			</tr>

		</table>
		</center>
		</font>

		<br>     
		<br>     

		<center>
		<form  METHOD="post"  NAME="UPLOAD1" ACTION="./index2.asp" >
			<input type="submit" value="戻る" id=submit4 name=submit4>
		</form>
		</center>

<br>     
<br>     
</body>     
</html>     
