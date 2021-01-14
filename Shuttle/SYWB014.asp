<%@ LANGUAGE="VBScript" %>
<!--#include file="Common.inc"-->
<html>

<head>
<title>搬出入予約申請作業詳細画面</title>
</head>

<body>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title02.gif" width="236" height="34"><p>
<b><font size=5>○○○○ユーザー向け画面　　○○○○グループ</font></b><br>
<b><font size=5>搬出入予約申請作業詳細</font>（2000.08.01 10:00現在）</b>
</center>
<br>   
<br>   
   
		<center>
		<table rules="none" width="600"  >   
					
			<tr>
				<td></td>
			    <td>＜申請内容＞</td>			
			    <td>＜受付結果＞</td>			
			</tr>

			<tr>
				<td>１．申請日</td>
				<td>2000/8/8</td>
				<td>2000/8/8</td>
			</tr>
				
			<tr>
				<td>２．作業種類</td>
				<td>搬出</td>
				<td>搬出</td>
			</tr>

			<tr>
				<td>３．コンテナ番号</td>
				<td>MESU00000006</td>
				<td>MESU00000006</td>
			</tr>


			<tr>
				<td>４．ＢＬ指定</td>
				<td>すべて</td>
				<td></td>
			</tr>


			<tr>
				<td>５．ブッキング番号</td>
				<td></td>
				<td></td>
			</tr>

			<tr>
				<td>６．搬出入日</td>
				<td>2000/08/10</td>
				<td>2000/08/10</td>
			</tr>

			<tr>
				<td>７．時間帯</td>
				<td>15:00まで</td>
				<td>14:00-15:00</td>
			</tr>

		</table>
		</center>


		<br>     
		<br>     

		<center>
		<table border="1" width="300"  >   
					
			<tr>
				<td bgcolor="#e8ffe8">作業番号</td>
			    <td bgcolor="#e8ffe8">状態</td>			
			    <td bgcolor="#e8ffe8">場所</td>			
			    <td bgcolor="#e8ffe8">シャーシＩＤ</td>			
			</tr>

			<tr>
				<td>10028</td>
			    <td>ダイヤ決</td>			
			    <td>ＣＹ</td>			
			    <td>　</td>			
			</tr>

		</table>
			</center>


		<br>     
		<br>     

		<center>
		<table border="0">   
		    <form  METHOD="post"  NAME="UPLOAD1" ACTION="./SYWB013.asp" >
				<td><input type="submit" value="戻る" id=submit4 name=submit4></td>
			</form>
		    <form  METHOD="post"  NAME="UPLOAD1" ACTION="./index2.asp" >
				<td><input type="submit" value="メニューへ" id=submit4 name=submit4></td>
			</form>
		    <form  METHOD="post"  NAME="UPLOAD1" ACTION="./SYWB013.asp" >
				<td><input type="submit" value="削除" id=submit4 name=submit4></td>
			</form>
		</table>
		</center>

<br>     
<br>     
</body>     
</html>     
