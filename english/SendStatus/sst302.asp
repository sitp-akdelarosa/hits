<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst302.asp				_/
'_/	Function	:ステータス配信対象項目登録・更新			_/
'_/	Date			:2004/01/05				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'''セッションの有効性をチェック
	CheckLoginH

	'''データ登録／更新しました画面にて「最新の情報に更新」でSubmitされた場合の対策
	'''まだデータ登録／更新しました画面は表示されていない場合
	if Session.Contents("ItemsSubmitted")="False" then

'''サーバ日付の取得
	Dim DayTime
	getDayTime DayTime

'''ユーザデータ所得
	Dim USER, sUN, Utype
	USER   = UCase(Session.Contents("userid"))
	sUN    = Session.Contents("sUN")
	Utype  = Session.Contents("UType")

'''データ取得
' 2009/03/10 R.Shibuta Upd-S
'	Dim F_ArrivalTime, F_InTime, F_List, F_DOStatus, F_DelPermit
'	Dim F_DemurrageFreeTime, F_CYDelTime, F_DetentionFreeTime, F_ReturnTime

	Dim F_ArrivalTime(4), F_InTime(4), F_List(4), F_DOStatus(4), F_DelPermit(4)
	Dim F_DemurrageFreeTime(4), F_CYDelTime(4), F_DetentionFreeTime(4), F_ReturnTime(4)

	Dim DaysToDMFT, DaysToDTFT
	Dim Email1, Email2, Email3, Email4, Email5
	Dim NewOrUpdate, tmpstr
	Dim iCnt

'	F_ArrivalTime = Request.Form("F_ArrivalTime")
'	F_InTime = Request.Form("F_InTime")
'	F_List = Request.Form("F_List")
'	F_DOStatus = Request.Form("F_DOStatus")
'	F_DelPermit = Request.Form("F_DelPermit")
'	F_DemurrageFreeTime = Request.Form("F_DemurrageFreeTime")
'	F_CYDelTime = Request.Form("F_CYDelTime")
'	F_DetentionFreeTime = Request.Form("F_DetentionFreeTime")
'	F_ReturnTime = Request.Form("F_ReturnTime")

	F_ArrivalTime(0) = Request.Form("F_ArrivalTime1")
	F_ArrivalTime(1) = Request.Form("F_ArrivalTime2")
	F_ArrivalTime(2) = Request.Form("F_ArrivalTime3")
	F_ArrivalTime(3) = Request.Form("F_ArrivalTime4")
	F_ArrivalTime(4) = Request.Form("F_ArrivalTime5")

	F_InTime(0) = Request.Form("F_InTime1")
	F_InTime(1) = Request.Form("F_InTime2")
	F_InTime(2) = Request.Form("F_InTime3")
	F_InTime(3) = Request.Form("F_InTime4")
	F_InTime(4) = Request.Form("F_InTime5")

	F_List(0) = Request.Form("F_List1")
	F_List(1) = Request.Form("F_List2")
	F_List(2) = Request.Form("F_List3")
	F_List(3) = Request.Form("F_List4")
	F_List(4) = Request.Form("F_List5")

	F_DOStatus(0) = Request.Form("F_DOStatus1")
	F_DOStatus(1) = Request.Form("F_DOStatus2")
	F_DOStatus(2) = Request.Form("F_DOStatus3")
	F_DOStatus(3) = Request.Form("F_DOStatus4")
	F_DOStatus(4) = Request.Form("F_DOStatus5")

	F_DelPermit(0) = Request.Form("F_DelPermit1")
	F_DelPermit(1) = Request.Form("F_DelPermit2")
	F_DelPermit(2) = Request.Form("F_DelPermit3")
	F_DelPermit(3) = Request.Form("F_DelPermit4")
	F_DelPermit(4) = Request.Form("F_DelPermit5")

	F_DemurrageFreeTime(0) = Request.Form("F_DemurrageFreeTime1")
	F_DemurrageFreeTime(1) = Request.Form("F_DemurrageFreeTime2")
	F_DemurrageFreeTime(2) = Request.Form("F_DemurrageFreeTime3")
	F_DemurrageFreeTime(3) = Request.Form("F_DemurrageFreeTime4")
	F_DemurrageFreeTime(4) = Request.Form("F_DemurrageFreeTime5")

	F_CYDelTime(0) = Request.Form("F_CYDelTime1")
	F_CYDelTime(1) = Request.Form("F_CYDelTime2")
	F_CYDelTime(2) = Request.Form("F_CYDelTime3")
	F_CYDelTime(3) = Request.Form("F_CYDelTime4")
	F_CYDelTime(4) = Request.Form("F_CYDelTime5")

	F_DetentionFreeTime(0) = Request.Form("F_DetentionFreeTime1")
	F_DetentionFreeTime(1) = Request.Form("F_DetentionFreeTime2")
	F_DetentionFreeTime(2) = Request.Form("F_DetentionFreeTime3")
	F_DetentionFreeTime(3) = Request.Form("F_DetentionFreeTime4")
	F_DetentionFreeTime(4) = Request.Form("F_DetentionFreeTime5")

	F_ReturnTime(0) = Request.Form("F_ReturnTime1")
	F_ReturnTime(1) = Request.Form("F_ReturnTime2")
	F_ReturnTime(2) = Request.Form("F_ReturnTime3")
	F_ReturnTime(3) = Request.Form("F_ReturnTime4")
	F_ReturnTime(4) = Request.Form("F_ReturnTime5")
' 2009/03/10 R.Shibuta Upd-E
	DaysToDMFT = Request.Form("DaysToDMFT")
	DaysToDTFT = Request.Form("DaysToDTFT")
	Email1 = Request.Form("Email1")
	Email2 = Request.Form("Email2")
	Email3 = Request.Form("Email3")
	Email4 = Request.Form("Email4")
	Email5 = Request.Form("Email5")

'エラートラップ開始
	on error resume next
	'''DB接続
	Dim ObjConn, ObjRS, StrSQL, RecordCNT
	ConnDBH ObjConn, ObjRS

	StrSQL = "SELECT count(*) AS NUM from TargetItems where UserCode='"& USER &"'"
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		Session.Contents("sst301") = "false"
		DisConnDBH ObjConn, ObjRS	'DB切断
		jumpErrorP "2","c103","10","ステータス配信対象項目登録・更新","101","SQL:<BR>"&strSQL
	else
		RecordCNT = ObjRS("NUM")
		ObjRS.close
	end if

	'''新規登録の場合
	if RecordCNT = 0 then
		StrSQL = "INSERT INTO TargetItems(UserCode,UpdtTime,UpdtPgCd,UpdtTmnl,"
		StrSQL = StrSQL & "FlagArrivalTime,FlagInTime,FlagList,FlagDOStatus,FlagDelPermit,FlagDemurrageFreeTime,"
		StrSQL = StrSQL & "DaysToDemurrageFreeTime,FlagCYDelTime,FlagDetentionFreeTime,DaysToDetentionFreeTime,"
		StrSQL = StrSQL & "FlagReturnTime,Email1,Email2,Email3,Email4,Email5, "
		StrSQL = StrSQL & "FlagArrivalTime2,FlagInTime2,FlagList2,FlagDOStatus2,FlagDelPermit2,FlagDemurrageFreeTime2,"
		StrSQL = StrSQL & "FlagCYDelTime2,FlagDetentionFreeTime2,FlagReturnTime2,"
		StrSQL = StrSQL & "FlagArrivalTime3,FlagInTime3,FlagList3,FlagDOStatus3,FlagDelPermit3,FlagDemurrageFreeTime3,"
		StrSQL = StrSQL & "FlagCYDelTime3,FlagDetentionFreeTime3,FlagReturnTime3,"
		StrSQL = StrSQL & "FlagArrivalTime4,FlagInTime4,FlagList4,FlagDOStatus4,FlagDelPermit4,FlagDemurrageFreeTime4,"
		StrSQL = StrSQL & "FlagCYDelTime4,FlagDetentionFreeTime4,FlagReturnTime4,"
' 2009/03/10 R.Shibuta Upd-S
		StrSQL = StrSQL & "FlagArrivalTime5,FlagInTime5,FlagList5,FlagDOStatus5,FlagDelPermit5,FlagDemurrageFreeTime5,"
		StrSQL = StrSQL & "FlagCYDelTime5,FlagDetentionFreeTime5,FlagReturnTime5) "
		StrSQL = StrSQL & "values('" & USER & "','" & Now() & "','STATUS01','" & USER & "',"
'		StrSQL = StrSQL & "'" & F_ArrivalTime & "','" & F_InTime & "','" & F_List & "',"
'		StrSQL = StrSQL & "'" & F_DOStatus & "','" & F_DelPermit & "',"

		StrSQL = StrSQL & "'" & F_ArrivalTime(0) & "','" & F_InTime(0) & "','" & F_List(0) & "',"
		StrSQL = StrSQL & "'" & F_DOStatus(0) & "','" & F_DelPermit(0) & "',"
' 2009/03/10 R.Shibuta Upd-E

' 2009/03/10 R.Shibuta Upd-S
'		if F_DemurrageFreeTime = "1" then
'			StrSQL = StrSQL & "'1'," & CInt(DaysToDMFT) & ","
'		else
'			StrSQL = StrSQL & "'0',NULL,"
'		end if

		StrSQL = StrSQL & "'" & F_DemurrageFreeTime(0) & "',"

		if F_DemurrageFreeTime(0) = "1" or _
			F_DemurrageFreeTime(1) = "1" or _
			F_DemurrageFreeTime(2) = "1" or _
			F_DemurrageFreeTime(3) = "1" or _
			F_DemurrageFreeTime(4) = "1" then
				StrSQL = StrSQL & CInt(DaysToDMFT) & ","
		else
				StrSQL = StrSQL & "NULL,"
		end if
' 2009/03/10 R.Shibuta Upd-E
		
		StrSQL = StrSQL & "'" & F_CYDelTime(0) & "',"

' 2009/03/10 R.Shibuta Upd-S
'		if F_DetentionFreeTime ="1" then
'			StrSQL = StrSQL & "'1'," & CInt(DaysToDTFT) & ","
'		else
'			StrSQL = StrSQL & "'0',NULL,"
'		end if


		StrSQL = StrSQL & "'" &  F_DetentionFreeTime(0) & "',"

		if F_DetentionFreeTime(0) = "1" or _
			F_DetentionFreeTime(1) = "1" or _
			F_DetentionFreeTime(2) = "1" or _
			F_DetentionFreeTime(3) = "1" or _
			F_DetentionFreeTime(4) = "1" then
				StrSQL = StrSQL & CInt(DaysToDTFT) & ","
		else
				StrSQL = StrSQL & "NULL,"
		end if
' 2009/03/10 R.Shibuta Upd-E

		StrSQL = StrSQL & "'" & F_ReturnTime(0) & "',"

		if Email1 <> "" then
			StrSQL = StrSQL & "'" & Email1 & "',"
		else
			StrSQL = StrSQL & "NULL,"
		end if
		if Email2 <> "" then
			StrSQL = StrSQL & "'" & Email2 & "',"
		else
			StrSQL = StrSQL & "NULL,"
		end if
		if Email3 <> "" then
			StrSQL = StrSQL & "'" & Email3 & "',"
		else
			StrSQL = StrSQL & "NULL,"
		end if
		if Email4 <> "" then
			StrSQL = StrSQL & "'" & Email4 & "',"
		else
			StrSQL = StrSQL & "NULL,"
		end if
		if Email5 <> "" then
			StrSQL = StrSQL & "'" & Email5 & "'"
		else
			StrSQL = StrSQL & "NULL"
		end if

' 2009/03/10 R.Shibuta Add-S
		For iCnt = 1 To 4
' 2009/03/10 R.Shibuta Add-E
			StrSQL = StrSQL & ","
			StrSQL = StrSQL & "'" & F_ArrivalTime(iCnt) & "','" & F_InTime(iCnt) & "','" & F_List(iCnt) & "',"
			StrSQL = StrSQL & "'" & F_DOStatus(iCnt) & "','" & F_DelPermit(iCnt) & "','" & F_DemurrageFreeTime(iCnt) & "',"
			StrSQL = StrSQL & "'" & F_CYDelTime(iCnt) & "','" & F_DetentionFreeTime(iCnt) & "','" & F_ReturnTime(iCnt) & "'"
' 2009/03/10 R.Shibuta Add-S
		Next
' 2009/03/10 R.Shibuta Add-E

		StrSQL = StrSQL & ")"
		
		ObjConn.Execute(StrSQL)
		if err <> 0 then
			Session.Contents("sst301") = "false"
			Set ObjRS = Nothing
			jumpErrorPDB ObjConn,"2","c104","10","ステータス配信対象項目登録","103","SQL:<BR>"&StrSQL
		end if

	'''更新の場合
	else
		StrSQL = "UPDATE TargetItems SET UpdtTime='"& Now() &"', UpdtPgCd='STATUS01', UpdtTmnl='"& USER &"',"
' 2009/03/10 R.Shibuta Upd-S
'		StrSQL = StrSQL & " FlagArrivalTime='" & F_ArrivalTime & "',"
'		StrSQL = StrSQL & " FlagInTime='" & F_InTime & "',"
'		StrSQL = StrSQL & " FlagList='" & F_List & "',"
'		StrSQL = StrSQL & " FlagDOStatus='" & F_DOStatus & "',"
'		StrSQL = StrSQL & " FlagDelPermit='" & F_DelPermit & "',"
'		StrSQL = StrSQL & " FlagDemurrageFreeTime='" & F_DemurrageFreeTime & "',"

		StrSQL = StrSQL & " FlagArrivalTime='" & F_ArrivalTime(0) & "',"
		StrSQL = StrSQL & " FlagInTime='" & F_InTime(0) & "',"
		StrSQL = StrSQL & " FlagList='" & F_List(0) & "',"
		StrSQL = StrSQL & " FlagDOStatus='" & F_DOStatus(0) & "',"
		StrSQL = StrSQL & " FlagDelPermit='" & F_DelPermit(0) & "',"
		StrSQL = StrSQL & " FlagDemurrageFreeTime='" & F_DemurrageFreeTime(0) & "',"

'		if F_DemurrageFreeTime = "1" then
'			StrSQL = StrSQL & " DaysToDemurrageFreeTime= " & CInt(DaysToDMFT) & ","
'		else
'			StrSQL = StrSQL & " DaysToDemurrageFreeTime=NULL,"
'		end if

		if F_DemurrageFreeTime(0) = "1" or _
		   F_DemurrageFreeTime(1) = "1" or _
		   F_DemurrageFreeTime(2) = "1" or _
		   F_DemurrageFreeTime(3) = "1" or _
		   F_DemurrageFreeTime(4) = "1" then
		 	StrSQL = StrSQL & " DaysToDemurrageFreeTime= '" & CInt(DaysToDMFT) & "',"
		else
		 	StrSQL = StrSQL & " DaysToDemurrageFreeTime=NULL,"
		end if
		 
		StrSQL = StrSQL & " FlagCYDelTime='" & F_CYDelTime(0) & "',"
		StrSQL = StrSQL & " FlagDetentionFreeTime='" & F_DetentionFreeTime(0) & "',"

'		if F_DetentionFreeTime = "1" then
'			StrSQL = StrSQL & " DaysToDetentionFreeTime=" & CInt(DaysToDTFT) & ","
'		else
'			StrSQL = StrSQL & " DaysToDetentionFreeTime=NULL,"
'		end if

		if F_DetentionFreeTime(0) = "1" or _
		   F_DetentionFreeTime(1) = "1" or _
		   F_DetentionFreeTime(2) = "1" or _
		   F_DetentionFreeTime(3) = "1" or _
		   F_DetentionFreeTime(4) = "1" then
		 	StrSQL = StrSQL & " DaysToDetentionFreeTime='" & CInt(DaysToDTFT) & "',"
		else
		 	StrSQL = StrSQL & " DaysToDetentionFreeTime=NULL,"
		end if
		
		StrSQL = StrSQL & " FlagReturnTime='" & F_ReturnTime(0) & "',"
' 2009/03/10 R.Shibuta Upd-E

		if Email1 <> "" then
			StrSQL = StrSQL & " Email1='" & Email1 & "',"
		else
			StrSQL = StrSQL & " Email1=NULL,"
		end if
		if Email2 <> "" then
			StrSQL = StrSQL & " Email2='" & Email2 & "',"
		else
			StrSQL = StrSQL & " Email2=NULL,"
		end if
		if Email3 <> "" then
			StrSQL = StrSQL & " Email3='" & Email3 & "',"
		else
			StrSQL = StrSQL & " Email3=NULL,"
		end if
		if Email4 <> "" then
			StrSQL = StrSQL & " Email4='" & Email4 & "',"
		else
			StrSQL = StrSQL & " Email4=NULL,"
		end if
		if Email5 <> "" then
			StrSQL = StrSQL & " Email5='" & Email5 & "',"
		else
			StrSQL = StrSQL & " Email5=NULL,"
		end if
' 2009/03/10 R.Shibuta Add-S
		StrSQL = StrSQL & " FlagArrivalTime2='" & F_ArrivalTime(1) & "',"
		StrSQL = StrSQL & " FlagInTime2='" & F_InTime(1) & "',"
		StrSQL = StrSQL & " FlagList2='" & F_List(1) & "',"
		StrSQL = StrSQL & " FlagDOStatus2='" & F_DOStatus(1) & "',"
		StrSQL = StrSQL & " FlagDelPermit2='" & F_DelPermit(1) & "',"
		StrSQL = StrSQL & " FlagDemurrageFreeTime2='" & F_DemurrageFreeTime(1) & "',"
		StrSQL = StrSQL & " FlagCYDelTime2='" & F_CYDelTime(1) & "',"
		StrSQL = StrSQL & " FlagDetentionFreeTime2='" & F_DetentionFreeTime(1) & "',"
		StrSQL = StrSQL & " FlagReturnTime2='" & F_ReturnTime(1) & "',"

		StrSQL = StrSQL & " FlagArrivalTime3='" & F_ArrivalTime(2) & "',"
		StrSQL = StrSQL & " FlagInTime3='" & F_InTime(2) & "',"
		StrSQL = StrSQL & " FlagList3='" & F_List(2) & "',"
		StrSQL = StrSQL & " FlagDOStatus3='" & F_DOStatus(2) & "',"
		StrSQL = StrSQL & " FlagDelPermit3='" & F_DelPermit(2) & "',"
		StrSQL = StrSQL & " FlagDemurrageFreeTime3='" & F_DemurrageFreeTime(2) & "',"
		StrSQL = StrSQL & " FlagCYDelTime3='" & F_CYDelTime(2) & "',"
		StrSQL = StrSQL & " FlagDetentionFreeTime3='" & F_DetentionFreeTime(2) & "',"
		StrSQL = StrSQL & " FlagReturnTime3='" & F_ReturnTime(2) & "',"

		StrSQL = StrSQL & " FlagArrivalTime4='" & F_ArrivalTime(3) & "',"
		StrSQL = StrSQL & " FlagInTime4='" & F_InTime(3) & "',"
		StrSQL = StrSQL & " FlagList4='" & F_List(3) & "',"
		StrSQL = StrSQL & " FlagDOStatus4='" & F_DOStatus(3) & "',"
		StrSQL = StrSQL & " FlagDelPermit4='" & F_DelPermit(3) & "',"
		StrSQL = StrSQL & " FlagDemurrageFreeTime4='" & F_DemurrageFreeTime(3) & "',"
		StrSQL = StrSQL & " FlagCYDelTime4='" & F_CYDelTime(3) & "',"
		StrSQL = StrSQL & " FlagDetentionFreeTime4='" & F_DetentionFreeTime(3) & "',"
		StrSQL = StrSQL & " FlagReturnTime4='" & F_ReturnTime(3) & "',"
	
		StrSQL = StrSQL & " FlagArrivalTime5='" & F_ArrivalTime(4) & "',"
		StrSQL = StrSQL & " FlagInTime5='" & F_InTime(4) & "',"
		StrSQL = StrSQL & " FlagList5='" & F_List(4) & "',"
		StrSQL = StrSQL & " FlagDOStatus5='" & F_DOStatus(4) & "',"
		StrSQL = StrSQL & " FlagDelPermit5='" & F_DelPermit(4) & "',"
		StrSQL = StrSQL & " FlagDemurrageFreeTime5='" & F_DemurrageFreeTime(4) & "',"
		StrSQL = StrSQL & " FlagCYDelTime5='" & F_CYDelTime(4) & "',"
		StrSQL = StrSQL & " FlagDetentionFreeTime5='" & F_DetentionFreeTime(4) & "',"
		StrSQL = StrSQL & " FlagReturnTime5='" & F_ReturnTime(4) & "'"	
' 2009/03/10 R.Shibuta Add-E
		StrSQL = StrSQL & " WHERE UserCode = '" & USER & "'"

		ObjConn.Execute(StrSQL)
		if err <> 0 then
			Session.Contents("sst301") = "false"
			Set ObjRS = Nothing
			jumpErrorPDB ObjConn,"2","c104","11","ステータス配信対象項目更新","103","SQL:<BR>"& err.description
		end if

	end if

'''DB接続解除
	DisConnDBH ObjConn, ObjRS
'''エラートラップ解除
	on error goto 0

	Session.Contents("sst301") = "false"


	'''ログ出力
' 2009/03/10 R.Shibuta upd-S
'	tmpstr = F_ArrivalTime & "," & F_InTime & "," & F_List & "," & F_DOStatus & "," &_
'						F_DelPermit & "," & F_DemurrageFreeTime & "," & DaysToDMFT & "," &_
'						F_CYDelTime & "," & F_DetentionFreeTime & "," & DaysToDTFT & "," &_
'						F_ReturnTime & "," & Email1 & "," & Email2 & "," & Email3 & "," & Email4 & "," & Email5

	tmpstr = F_ArrivalTime(0) & "," & F_InTime(0) & "," & F_List(0) & "," & F_DOStatus(0) & "," &_
						F_DelPermit(0) & "," & F_DemurrageFreeTime(0) & "," & DaysToDMFT & "," &_
						F_CYDelTime(0) & "," & F_DetentionFreeTime(0) & "," & DaysToDTFT & "," &_
						F_ReturnTime(0) & "," & Email1 & "," & Email2 & "," & Email3 & "," & Email4 & "," & Email5
						
	For iCnt = 1 To 4
		tmpStr = tmpStr & "," & _
		F_ArrivalTime(iCnt) & "," & _ 
		F_InTime(iCnt) & "," & F_List(iCnt) & "," & F_DOStatus(iCnt) & "," &_
		F_DelPermit(iCnt) & "," & F_DemurrageFreeTime(iCnt) & "," & _
		F_CYDelTime(iCnt) & "," & F_DetentionFreeTime(iCnt) & "," & _
		F_ReturnTime(iCnt)
	Next
' 2009/03/10 R.Shibuta upd-E
	if RecordCNT = 0 then
		WriteLogH "c103", "ステータス配信対象設定","10",tmpstr
	else
		WriteLogH "c103", "ステータス配信対象設定","11",tmpstr
	end if

	Session.Contents("ItemsSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信対象項目登録・更新</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信対象項目登録・更新--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
	<TR>
		<TD align=center>
<% if RecordCNT = "0" then %>
			<BR><BR>登録しました。<BR><BR><BR>
<% end if %>
<% if RecordCNT = "1" then %>
			<BR><BR>更新しました。<BR><BR><BR>
<% end if %>
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("ItemsSubmitted")="False"のelse処理 %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信対象項目登録・更新</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信対象項目登録・更新--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
	<TR>
		<TD align=center>
			<BR><BR>登録・更新は既に完了しています。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("ItemsSubmitted")="False"のendif処理 %>
<% end if %>
