<!-- #include file="Common.inc" -->

<html>
<head>
<title> バナーアクセス数集計 </title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>


<%
Function MonthName(inputM)
Select case inputM '1～12の数を入れると英語で月名を返す
case 1
	MonthName = "January"
case 2
	MonthName = "February"
case 3
	MonthName = "March"
case 4
	MonthName = "April"
case 5
	MonthName = "May"
case 6
	MonthName = "June"
case 7
	MonthName = "July"
case 8
	MonthName = "August"
case 9
	MonthName = "September"
case 10
	MonthName = "October"
case 11
	MonthName = "November"
case 12
	MonthName = "December"
End Select
End Function


Dim error		'引数エラーフラグ
Dim aMonth		'accessのあった年月(int yyyymm)
Dim cMonth		'count upする月の英名(varchar)
Dim bNumber		'バナーの位置(int 左から1,2,3…)
Dim lUrl		'広告のURL
aMonth = Year(Date)*100 + Month(Date)
cMonth = MonthName(aMonth mod 100)
bNumber = request.querystring("banner_number")
lUrl = request.querystring("link_url")

'index側からの引数が不正な時のエラー処理
If bNumber = "" then
	response.write("バナーNoが取得できません｡<BR>")
	error = true
End if
If lUrl = "" then
	response.write("URLが取得できません。<BR>")
	error = true
End if
If error then
	response.write("indexを確認して下さい。")
	response.end
End if

ConnectSvr conn, rs
sql = "SELECT * FROM banner_click "
sql = sql& " WHERE banner_no=" &bNumber& " AND last_clicked_date=" &aMonth 
rs.Open sql, conn
If rs.EOF then 			'クリック年月≠最終クリック年月の場合はゼロクリア処理
'	response.write("今月のアクセスが確認できません。<BR>")

	rs.Close
	sql = "SELECT * FROM banner_click "
	sql = sql& " WHERE banner_no=" &bNumber
	rs.Open sql, conn

	If IsNull(rs("last_clicked_date")) then '新広告追加時のゼロクリア(テーブルは手動で追加が前提)
'		response.write("新しい広告のテーブルを初期化します。<BR>")
		For i = 1 To 12
			sql = "UPDATE banner_click SET " &MonthName(i)& "= 0"
			sql = sql & " WHERE banner_no =" &bNumber
		conn.execute sql
		Next
	Else		'単なる月初めの場合のゼロクリア
'		response.write("今月初めてのアクセスのようです｡テーブルを書き換えます｡<BR>")
		For i = rs("last_clicked_date")+1 To aMonth
			If i mod 100 > 12 then	'yyyymmからmmを取り出す時の、
				i = i + 100 - 12	'for文によるインクリメントの影響を打ち消す
			End if
			sql = "UPDATE banner_click SET " &MonthName(i mod 100)& "= 0"
			sql = sql & " WHERE banner_no =" &bNumber
			conn.execute sql
		Next
	End if
End if


'アクセス数のインクリメント、各種データの更新
sql = "UPDATE banner_click SET " &cMonth& " = " &cMonth& "+ 1, UpdtPgCd='banner', UpdtTmnl='banner', last_clicked_date=" &aMonth& ", UpdtTime='" &Now()& "'"
sql = sql & " WHERE banner_no =" &bNumber
conn.execute sql

'終了処理
rs.Close
response.redirect(lUrl)
'response.write("すべての処理が終了しました｡")
%>


<body>
</body>
</html>
