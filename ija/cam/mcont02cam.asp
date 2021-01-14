<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common_cam.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim vCtno, vCtnoE, vCtnoS
Dim CntNo,sCntNo2
Dim vFlg								'コンテナ照会より遷移("1")／ＢＬ照会遷移(△)
Dim sql
Dim sErrMsg
Dim sErrOpt

sErrMSg = ""
sErrOpt = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

vCtno = Trim(Request.QueryString("Ctno"))
vCtnoE = Trim(Request.QueryString("cont_e"))
vCtnoS = Trim(Request.QueryString("cont_s"))
If IsEmpty(vCtno) Or vCtno ="" Then
	vFlg = "1"
	vCtno = Ucase(vCtnoE) & vCtnoS
End If

Dim conn, rs
ConnectSvr conn, rs

CntNo = vCtno
If CntNo = "" Then
	sErrMsg = "ｺﾝﾃﾅ未入力"
Else
	'コンテナ番号の数値部分のみ入力されている場合、該当するコンテナを探す
	If vFlg = "1" And  (vCtnoE = "" Or IsEmpty(vCtnoE)) Then	
		sql = "SELECT RTrim([ContNo]) AS CT  FROM ImportCont GROUP BY RTrim([ContNo]), ContNo "
		sql = sql  & "HAVING (((RTrim([ContNo])) Like '%" & vCtnoS & "'))"
		rs.Open sql, conn, 0, 1, 1
		If rs.Eof Then
			sErrMsg = "該当コンテナなし"
			sErrOpt = vCtnoS
		Else
			CntNo = rs("CT")		'コンテナ番号再設定
			rs.MoveNext
			Do While Not rs.EOF
				sCntNo2 = rs("CT")
				rs.MoveNext
				If CntNo<>sCntNo2 Then
					sErrMsg = "ｺﾝﾃﾅ複数存在"
					sErrOpt = vCtnoS
					Exit Do
				End If
			Loop
		End If
		rs.Close
	End If
End If

If sErrMSg = "" Then
'--- mod by MES(2004/9/10)
'	sql = "SELECT ImportCont.ContNo, ImportCont.DGFlag, ImportCont.WHArSchDate, ImportCont.RFFlag, " & _
'	      " ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.CYDelTime, " & _
'	      " ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, " & _
'	      " ImportCont.FreeTimeExt, Container.ContSize, Container.ContHeight, " & _
'		  " BL.RecTerminal, BL.RFFlag BRFFlag, BL.DGFlag BDGFlag " & _
'		  " FROM ImportCont, Container, BL " & _
'		  " WHERE Container.ContNo='" & CntNo & "' " & _
'		  " AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
'	      " AND Container.ContNo=ImportCont.ContNo " & _
'	      " AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl " & _
'	      " AND BL.BLNo=*ImportCont.BLNo"
'--- mod by MES(2005/3/28)
'	sql = "SELECT ImportCont.ContNo, ImportCont.DGFlag, ImportCont.WHArSchDate, ImportCont.RFFlag, " & _
'	      " ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.OLTICDate, ImportCont.CYDelTime, " & _
'	      " ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, " & _
'	      " ImportCont.FreeTimeExt, Container.ContSize, Container.ContHeight, " & _
'		  " BL.RecTerminal, BL.RFFlag BRFFlag, BL.DGFlag BDGFlag " & _
'		  " FROM ImportCont, Container, BL " & _
'		  " WHERE Container.ContNo='" & CntNo & "' " & _
'		  " AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
'	      " AND Container.ContNo=ImportCont.ContNo " & _
'	      " AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl " & _
'	      " AND BL.BLNo=*ImportCont.BLNo"
	sql = "SELECT ImportCont.ContNo, ImportCont.DGFlag, ImportCont.WHArSchDate, ImportCont.RFFlag, " & _
	      " ImportCont.FreeTime, ImportCont.OLTICFlag, ImportCont.OLTICNo, ImportCont.OLTICDate, ImportCont.CYDelTime, " & _
	      " ImportCont.DOStatus, ImportCont.DelPermitDate, ImportCont.OLTDateFrom, ImportCont.OLTDateTo, " & _
	      " ImportCont.FreeTimeExt, Container.ContSize, Container.ContHeight, " & _
	      " Container.ListNo, Container.OffDockFlag, Container.DsListFlg, " & _
		  " BL.RecTerminal, BL.RFFlag BRFFlag, BL.DGFlag BDGFlag " & _
		  " FROM ImportCont, Container, BL " & _
		  " WHERE Container.ContNo='" & CntNo & "' " & _
		  " AND Container.VslCode=ImportCont.VslCode AND Container.VoyCtrl=ImportCont.VoyCtrl " & _
	      " AND Container.ContNo=ImportCont.ContNo " & _
	      " AND BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl " & _
	      " AND BL.BLNo=*ImportCont.BLNo"
'--- end MES
'--- end MES
	rs.Open sql, conn, 0, 1, 1
	If rs.eof Then
		sErrMsg = "該当コンテナなし"
		sErrOpt = CntNo
	Else
		' 場所／コンテナサイズ
		Dim sPlace
		sPlace = Trim(rs("RecTerminal")) & "／" & Trim(rs("ContSize")) & "ft"

		' 危険物
		Dim sDanger
		sDanger=rs("DGFlag")
		If IsNull(sDanger) Or sDanger="" Then
			sDanger=rs("BDGFlag")
		End If
		If sDanger = "H" Then
'☆☆☆ Mod_S  by nics 2009.03.17
'			sDanger = "危険物:○"
'		Else
'			sDanger = "危険物:－"
			sDanger = "危険物品:○"
		Else
			sDanger = "危険物品:－"
'☆☆☆ Mod_E  by nics 2009.03.17
		End If

		' 倉庫到着指示時刻
		Dim sArriveTime, sYear, sMonth, sDay, sHour, sMinute
		sArriveTime = "倉庫到着指示時刻<br>　"
		If Not IsNull(rs("WHArSchDate")) Then
			sYear = CStr(Year(rs("WHArSchDate")))
			sMonth = Right(CStr(Month(rs("WHArSchDate")) + 100), 2)
			sDay = Right(CStr(Day(rs("WHArSchDate")) + 100), 2)
			sHour = Right(CStr(Hour(rs("WHArSchDate")) + 100), 2)
			sMinute = Right(CStr(Minute(rs("WHArSchDate")) + 100), 2)
			sArriveTime = sArriveTime & sYear & "/" & sMonth & "/" & sDay & "　"  & sHour & ":" & sMinute
		End If

		' 高さ
		Dim sHeight
		sHeight = "高さ:" & Trim(rs("ContHeight"))

		' リーファー
		Dim sReefer
		sReefer = rs("RFFlag")
		If IsNull(sReefer) Or sReefer="" Then
			sReefer=rs("BRFFlag")
		End If
		If sReefer = "R" Then
			sReefer = "リーファー:○"
		Else
			sReefer = "リーファー:－"
		End If

'☆☆☆ Add_S  by nics 2009.03.17
		' フリータイム
		Dim sFreeTime
		sFreeTime = "フリータイム:" 
		If Not IsNull(rs("FreeTimeExt")) Then
			sMonth = Right(CStr(Month(rs("FreeTimeExt")) + 100), 2)
			sDay = Right(CStr(Day(rs("FreeTimeExt")) + 100), 2)
			sFreeTime = sFreeTime & sMonth & "/" & sDay 
		ElseIf Not IsNull(rs("FreeTime")) Then
			sMonth = Right(CStr(Month(rs("FreeTime")) + 100), 2)
			sDay = Right(CStr(Day(rs("FreeTime")) + 100), 2)
			sFreeTime = sFreeTime & sMonth & "/" & sDay 
		End If

		' 通関
		Dim strTsukan, strchkNow, strchkOLTDateFrom, strchkOLTDateTo
        strchkNow = DispDateTime( Now, 8 )
        strchkOLTDateFrom = DispDateTime( rs("OLTDateFrom"), 8 )
        strchkOLTDateTo = DispDateTime( rs("OLTDateTo"), 8 )
        ' オンドックで卸リスト対象外フラグが対象外でないなら卸リスト番号の有無をチェック
        If Trim(rs("OffDockFlag"))="N" And (Trim(rs("DsListFlg"))<>"1" Or IsNull(Trim(rs("DsListFlg"))) = True ) Then
        	If Trim(rs("ListNo"))="" Then
        		strTsukan = ""
        	End If
        End If
        ' 通関／保税輸送のチェック
        If Trim(rs("OLTICFlag"))="I" Then
            If Trim(rs("OLTICNo"))<>"" Then
                strTsukan = "I"
            Else
                strTsukan = ""
            End If
		ElseIf Trim(rs("OLTICFlag"))<>"" Then
		'OLTICFlagが空白でないとき、日付チェックを行う
            If strchkNow>=strchkOLTDateFrom And strchkNow<=strchkOLTDateTo Then
                strTsukan = "O"
            Else
                strTsukan = ""
            End If
		'OLTICFlagが空白のとき、許可日と許可番号のチェックをし、通関OKとする
        Else
			If DispDateTime(rs("OLTICDate"),8)<>"" AND Trim(rs("OLTICNo"))<>"" Then
				strTsukan = "N"
			End If
        End If
        If DispDateTime(rs("CYDelTime"),0)<>"" Then           ' 搬出されていたら○とする
            If IsNull(rs("OLTDateFrom")) Or IsNull(rs("OLTDateTo")) Then
                strTsukan = "S"
            Else
                strTsukan = "T"
            End If
        End If
		If strTsukan <> "" Then
			strTsukan = "通関:○"
		Else
			strTsukan = "通関:×"
		End If

		' D/O
		Dim sDOStatus
		sDOStatus = "D/O:"
		If rs("DOStatus") <> "Y" Then
			sDOStatus = sDOStatus & "×"
		Else
			sDOStatus = sDOStatus & "○"
		End If
'☆☆☆ Add_E  by nics 2009.03.17

		' 搬出可能か
		Dim sCarryOut, sCarryOutFlg
		Do While Not rs.Eof
			sCarryOutFlg = CanCarryOut(rs)
			If sCarryOutFlg<>" " Then
				If sCarryOutFlg="Y" Then
					sCarryOut = "搬出：○"
				Else
					sCarryOut = "搬出：済"
				End If
				rs.MoveNext
			Else
				sCarryOut = "搬出：×"
				Exit Do
			End If
		Loop
	End If
	rs.Close
End If
conn.Close

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If vFlg="1" Then
	If sErrMsg<>"" Then
		WriteLogM oFS, "Unknown", "2401", "携帯-コンテナ番号照会（中央ふ頭）", "10",sPhoneType, Ucase(vCtnoE) & "/" & vCtnoS & "," & "入力内容の正誤:1(誤り)" & sErrMsg
	Else
		WriteLogM oFS, "Unknown", "2401", "携帯-コンテナ番号照会（中央ふ頭）", "10",sPhoneType, Ucase(vCtnoE) & "/" & vCtnoS & "," & "入力内容の正誤:0(正しい)"
		WriteLogM oFS, "Unknown", "2402", "携帯-コンテナ詳細（中央ふ頭）", "00",sPhoneType, CntNo & ","
	End If
Else
	WriteLogM oFS, "Unknown", "2405", "携帯-コンテナ詳細(BL)（中央ふ頭）", "00",sPhoneType, CntNo & ","
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="コンテナ番号照会">
	<center>
	【ｺﾝﾃﾅ番号照会】<br><br>
<%
	If sErrMsg <> "" Then
		If sErrOpt <> "" Then
%>
			<center>
			<%=sErrOpt%><br>
<%
		End If
%>
		<center>
		<%=sErrMsg%><br>
<%
	Else
%>
		<center>
		<%=CntNo%><br>
		<center>
		<%=sCarryOut%><br>
		<center>
		<%=sPlace%><br>
		<center>
		---(以下詳細)---<br>
		<%=sHeight%><br>
		<%=sReefer%><br>
		<%=sDanger%><br>
<!-- add by nics 2009.03.17 -->
		<%=strTsukan%><br>
		<%=sDOStatus%><br>
		<%=sFreeTime%><br>
<!-- end of add by nics 2009.03.17 -->
		<%=sArriveTime%><br>
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
		<%=GetTitleTag("コンテナ番号照会")%>
	</head>
	
	<body>
	<center>
	【ｺﾝﾃﾅ番号照会】
	<hr>
<%
	If sErrMsg <> "" Then
		If sErrOpt <> "" Then
%>
			<%=sErrOpt%><br>
<%
		End If
%>
		<%=sErrMsg%><br><br>
<%
	Else
%>
		<%=CntNo%><br>
		<%=sCarryOut%><br>
		<%=sPlace%><br>
		---(以下詳細)---<br>
<%
		If sPhoneType <> "P" Then
			'PC以外は左詰(PCは画面が広すぎるので左詰めしない)
%>
			</center>
<%
		End If
%>
		<%=sHeight%><br>
		<%=sReefer%><br>
		<%=sDanger%><br>
<!-- add by nics 2009.03.17 -->
		<%=strTsukan%><br>
		<%=sDOStatus%><br>
		<%=sFreeTime%><br>
<!-- end of add by nics 2009.03.17 -->
		<%=sArriveTime%><br>
		<center>
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
