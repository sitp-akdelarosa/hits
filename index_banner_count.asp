<!-- #include file="Common.inc" -->

<html>
<head>
<title> �o�i�[�A�N�Z�X���W�v </title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>


<%
Function MonthName(inputM)
Select case inputM '1�`12�̐�������Ɖp��Ō�����Ԃ�
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


Dim error		'�����G���[�t���O
Dim aMonth		'access�̂������N��(int yyyymm)
Dim cMonth		'count up���錎�̉p��(varchar)
Dim bNumber		'�o�i�[�̈ʒu(int ������1,2,3�c)
Dim lUrl		'�L����URL
aMonth = Year(Date)*100 + Month(Date)
cMonth = MonthName(aMonth mod 100)
bNumber = request.querystring("banner_number")
lUrl = request.querystring("link_url")

'index������̈������s���Ȏ��̃G���[����
If bNumber = "" then
	response.write("�o�i�[No���擾�ł��܂���<BR>")
	error = true
End if
If lUrl = "" then
	response.write("URL���擾�ł��܂���B<BR>")
	error = true
End if
If error then
	response.write("index���m�F���ĉ������B")
	response.end
End if

ConnectSvr conn, rs
sql = "SELECT * FROM banner_click "
sql = sql& " WHERE banner_no=" &bNumber& " AND last_clicked_date=" &aMonth 
rs.Open sql, conn
If rs.EOF then 			'�N���b�N�N�����ŏI�N���b�N�N���̏ꍇ�̓[���N���A����
'	response.write("�����̃A�N�Z�X���m�F�ł��܂���B<BR>")

	rs.Close
	sql = "SELECT * FROM banner_click "
	sql = sql& " WHERE banner_no=" &bNumber
	rs.Open sql, conn

	If IsNull(rs("last_clicked_date")) then '�V�L���ǉ����̃[���N���A(�e�[�u���͎蓮�Œǉ����O��)
'		response.write("�V�����L���̃e�[�u�������������܂��B<BR>")
		For i = 1 To 12
			sql = "UPDATE banner_click SET " &MonthName(i)& "= 0"
			sql = sql & " WHERE banner_no =" &bNumber
		conn.execute sql
		Next
	Else		'�P�Ȃ錎���߂̏ꍇ�̃[���N���A
'		response.write("�������߂ẴA�N�Z�X�̂悤�ł���e�[�u�������������܂��<BR>")
		For i = rs("last_clicked_date")+1 To aMonth
			If i mod 100 > 12 then	'yyyymm����mm�����o�����́A
				i = i + 100 - 12	'for���ɂ��C���N�������g�̉e����ł�����
			End if
			sql = "UPDATE banner_click SET " &MonthName(i mod 100)& "= 0"
			sql = sql & " WHERE banner_no =" &bNumber
			conn.execute sql
		Next
	End if
End if


'�A�N�Z�X���̃C���N�������g�A�e��f�[�^�̍X�V
sql = "UPDATE banner_click SET " &cMonth& " = " &cMonth& "+ 1, UpdtPgCd='banner', UpdtTmnl='banner', last_clicked_date=" &aMonth& ", UpdtTime='" &Now()& "'"
sql = sql & " WHERE banner_no =" &bNumber
conn.execute sql

'�I������
rs.Close
response.redirect(lUrl)
'response.write("���ׂĂ̏������I�����܂����")
%>


<body>
</body>
</html>
