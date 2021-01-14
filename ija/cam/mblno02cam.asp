<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common_cam.inc"-->
<!--#include file="mcommon.inc"-->
<%
Const PAGE_SIZE = 10
Dim vPageNo
Dim nPageNo
Dim nPageMax
Dim vBlno
Dim sBlno
Dim sql
Dim sErrMsg
Dim sErrOpt
Dim iRecCnt
Dim sContNo(20)
Dim nContCount
Dim nContPtr

sErrMsg = ""
sErrOpt = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

vPageNo = Request.QueryString("PAGENO")
vBlno = Request.QueryString("BLno")

If IsEmpty(vPageNo) Or vPageNo ="" Then
	vPageNo = "1"
	vBlno = Ucase(Trim(Request.QueryString("BLno")))
End If
nPageNo = CInt(vPageNo)

Dim conn, rs
ConnectSvr conn, rs

iRecCnt = 0
If vBlno = "" Then
	sErrMsg = "Ｂ／Ｌ未入力"
Else
	sBlno = Trim(vBlno)

	' ＢＬ番号like検索
	If Len(sBlno) <= 19 Then
		dim iblcnt
		dim slblno

		iblcnt = 0
		slblno = "%" & sBlno
		sql = "SELECT RTrim([BLNo]) AS BLN FROM ImportCont GROUP BY RTrim([BLNo]), BLNo "
		sql = sql  & "HAVING (((RTrim([BLNo])) Like '" & slblno & "'))"
		rs.Open sql, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "該当Ｂ／Ｌなし"
			sErrOpt = Trim(vBlno)
		Else
			sBlno = rs("BLN")		'ＢＬ番号再設定
			rs.MoveNext
			If Not rs.Eof Then
				sErrMsg = "BL複数存在します"
				sErrOpt = Trim(vBlno)
			End If
		End If
		rs.Close
	End If
End If

If sErrMsg = "" Then
'--- mod by MES(2004/9/10)
'	sql = "SELECT ContNo, FreeTime, OLTICFlag, OLTICNo, CYDelTime, " & _
'		" DOStatus, DelPermitDate, OLTDateFrom, OLTDateTo, FreeTimeExt " & _
'		" FROM ImportCont WHERE BLNo='" & sBlno & "' " & _
'		" ORDER BY ContNo"
'--- mod by MES(2005/3/28)
'	sql = "SELECT ContNo, FreeTime, OLTICFlag, OLTICNo, OLTICDate, CYDelTime, " & _
'		" DOStatus, DelPermitDate, OLTDateFrom, OLTDateTo, FreeTimeExt " & _
'		" FROM ImportCont WHERE BLNo='" & sBlno & "' " & _
'		" ORDER BY ContNo"
	sql = "SELECT ImportCont.ContNo, ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, " & _
		" ImportCont.OLTICDate, ImportCont.CYDelTime, ImportCont.DOStatus, " & _
		" ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, ImportCont.FreeTimeExt, " & _
		" Container.ListNo, Container.OffDockFlag, Container.DsListFlg " & _
		" FROM ImportCont, Container WHERE ImportCont.BLNo='" & sBlno & "' " & _
		" AND Container.ContNo=ImportCont.ContNo AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
		" ORDER BY Container.ContNo"
'--- end MES
'--- end MES
	rs.Open sql, conn, 0, 1, 1
	If rs.eof Then
		sErrMsg = "該当Ｂ／Ｌなし"
		sErrOpt = Trim(vBlno)
	Else
		nContCount = 0
		Do While Not rs.Eof
			'搬出可能
			If CanCarryOut(rs)="Y" Then
				iRecCnt = iRecCnt + 1							'実際データカウント
				If (nPageNo - 1) * PAGE_SIZE < iRecCnt And iRecCnt <= nPageNo * PAGE_SIZE Then
					nContCount = nContCount + 1
					sContNo(nContCount) = Trim(rs("ContNo"))
				End If
			End If
			rs.MoveNext
		loop

		If iRecCnt =  0 Then
			sErrOpt = "搬出可能"
			sErrMsg = "コンテナ無し"
		Else
			'全ページ数
			nPageMax = -Int(-iRecCnt / PAGE_SIZE)
		End If
	End If
	rs.Close
End If
conn.Close

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If sErrMsg<>"" Then
	WriteLogM oFS, "Unknown", "2403", "携帯-BL番号照会（中央ふ頭）", "10",sPhoneType, sBLNo & "," & "入力内容の正誤:1(誤り)" & sErrMsg
Else
	WriteLogM oFS, "Unknown", "2403", "携帯-BL番号照会（中央ふ頭）", "10",sPhoneType, sBLNo & "," & "入力内容の正誤:0(正しい)"
	WriteLogM oFS, "Unknown", "2404", "携帯-コンテナ番号一覧（中央ふ頭）", "00",sPhoneType, nPageNo & "/" & nPageMax & ","
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="BL番号照会">
	<center>
	【BL番号照会】<br>
<%
	If sErrMsg <> "" Then
%>
		<center>
		<%=sErrOpt%><br>
		<center>
		<%=sErrMsg%><br><br>
<%
	Else
		For nContPtr = 1 To nContCount
%>
			<center>
			<a task="gosub" accesskey="<%=CStr(nContPtr Mod 10)%>"
						dest="mcont02cam.asp?Ctno=<%=sContNo(nContPtr)%>">
				<%=sContNo(nContPtr)%>
			</a><br>
<%
		Next
%>
		<center>
		-<%=nPageNo%>/<%=nPageMax%>-<br>
		<center>
<%
		If 1 < nPageNo Then
%>
			<a task="gosub" dest="mblno02cam.asp?PAGENO=<%=nPageNo - 1%>&BLno=<%=sBlno%>">前頁</a>
<%
		End If
		If nPageNo < nPageMax Then
%>
			<a task="gosub" dest="mblno02cam.asp?PAGENO=<%=nPageNo + 1%>&BLno=<%=sBlno%>">次頁</a>
<%
		End If
%>
		<br>
<%
	End If
%>
	<center>
	<a task="gosub" dest="index.asp">ﾒﾆｭｰ</a>
	</display>
	</hdml>
<%
Else
	' EzWeb以外のタグを編集
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("BL番号照会")%>
	</head>
	<body>
	<center>
	【BL番号照会】
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<%=sErrOpt%><br>
		<%=sErrMsg%><br><br>
<%
	Else
		For nContPtr = 1 To nContCount
%>
			<a href="mcont02cam.asp?Ctno=<%=sContNo(nContPtr)%>" <%=GetKeyTag(CStr(nContPtr))%>>
				<%=GetKeyLabel(CStr(nContPtr))%><%=sContNo(nContPtr)%>
			</a><br>
<%
		Next
%>
		-<%=nPageNo%>/<%=nPageMax%>-<br>
<%
		If 1 < nPageNo Then
%>
			<a href="mblno02cam.asp?PAGENO=<%=nPageNo - 1%>&BLno=<%=sBlno%>">前頁</a>
<%
		End If
		If nPageNo < nPageMax Then
%>
			<a href="mblno02cam.asp?PAGENO=<%=nPageNo + 1%>&BLno=<%=sBlno%>">次頁</a>
<%
		End If
%>
		<br>
<%
	End If
%>
	<form action="../index.asp" method="get">
		<input type="submit" value="ﾒﾆｭｰ">
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
