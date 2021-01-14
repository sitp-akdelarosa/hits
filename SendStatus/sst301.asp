<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst301.asp				_/
'_/	Function	:ステータス配信対象項目設定確認画面			_/
'_/	Date			:2003/12/27				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
	CheckLoginH

'データ取得
'2009/03/10 R.Shibuta Upd-S
'	Dim F_ArrivalTime, F_InTime, F_List, F_DOStatus, F_DelPermit
'	Dim F_DemurrageFreeTime, F_CYDelTime, F_DetentionFreeTime, F_ReturnTime
	Dim F_ArrivalTime(4), F_InTime(4), F_List(4), F_DOStatus(4), F_DelPermit(4)
	Dim F_DemurrageFreeTime(4), F_CYDelTime(4), F_DetentionFreeTime(4), F_ReturnTime(4)
	Dim DaysToDMFT, DaysToDTFT
	Dim Email1, Email2, Email3, Email4, Email5
	Dim NoEntered, ItemsToSend
	Dim iCnt
	DIm strWork
	Dim iChk

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

	DaysToDMFT = Request.Form("DaysToDemurrageFreeTime")
	DaysToDTFT = Request.Form("DaysToDetentionFreeTime")
	Email1 = Trim(Request.Form("Email1"))
	Email2 = Trim(Request.Form("Email2"))
	Email3 = Trim(Request.Form("Email3"))
	Email4 = Trim(Request.Form("Email4"))
	Email5 = Trim(Request.Form("Email5"))
'2009/03/10 R.Shibuta Upd-E

	'エラートラップ開始
'	on error resume next
	'DB接続
'	Dim ObjConn, ObjRS, StrSQL
'	ConnDBH ObjConn, ObjRS
	'DB接続解除
'	DisConnDBH ObjConn, ObjRS
	'エラートラップ解除
'	on error goto 0

	Session.Contents("sst301") = "true"

'2009/03/10 R.Shibuta Upd-S
	'''何も入力されていない場合
'	if F_ArrivalTime = "0" and F_InTime = "0" and F_List = "0" and F_DOStatus = "0" _
'		and F_DelPermit = "0" and F_DemurrageFreeTime = "0" and F_CYDelTime = "0" _
'		and F_DetentionFreeTime = "0" and F_ReturnTime = "0" _
'		and Email1 = "" and Email2 = "" and Email3 = "" and Email4 = "" and Email5 = "" then
'		NoEntered = "1"
'	else
'		NoEntered = "0"
'	end if

	For iCnt = 0 To 4
		if F_ArrivalTime(iCnt) = "0" and F_InTime(iCnt) = "0" and F_List(iCnt) = "0" and F_DOStatus(iCnt) = "0" _
			and F_DelPermit(iCnt) = "0" and F_DemurrageFreeTime(iCnt) = "0" and F_CYDelTime(iCnt) = "0" _
			and F_DetentionFreeTime(iCnt) = "0" and F_ReturnTime(iCnt) = "0" _
			and Email1 = "" and Email2 = "" and Email3 = "" and Email4 = "" and Email5 = "" then
			NoEntered = "1"
		else
			NoEntered = "0"
		end if
		
		if NoEntered = "0" then
			Exit For
		end if
	Next
'2009/03/10 R.Shibuta Upd-E

	'''メール送信対象項目数
	ItemsToSend = 0

	'''ログ出力
	WriteLogH "c103", "ステータス配信対象設定","02",""

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信対象項目設定入力確認</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(730,650);

window.focus();

//登録
function GoEntry(){
	f=document.sst301;
	f.action="sst302.asp";
	return true;
}
//戻る
function GoBack(){
	f=document.sst301;
	f.action="sst300.asp";
	return true;
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信対象項目設定入力確認画面--------------------------->
<% if NoEntered = "0" then %>
<FORM name="sst301" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
	<TR>
		<TD colspan="12">
			<B>輸入ステータス配信依頼（設定）項目確認</B>
		</TD>
	</TR>
	<TR><TD colspan="12">　</TD></TR>
	<TR>
		<TD colspan="12">
			以下の値が追加、または、変化する度に情報を送信します。
		</TD>
	</TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_ArrivalTime(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(1)入港時間</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_InTime(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(2)ＣＹ搬入時間</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_List(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(3)通関許可状況</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>	
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_DOStatus(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(4)ＤＯクリア状況</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_DelPermit(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(5)搬出可否</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>	
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_DemurrageFreeTime(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="50%"colspan="2">(6)デマレージフリータイム</TD>
			<TD width="45%"colspan="9"><%=DaysToDMFT%>日以内になったとき</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_CYDelTime(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(7)ＣＹ搬出時間</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_DetentionFreeTime(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="50%" colspan="2">(8)ディテンションフリータイム</TD>
			<TD width="45%" colspan = "9"><%=DaysToDTFT%>日以内になったとき</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<!-- 2009/03/10 R.Shibuta Add-S -->
<% For iCnt = 0 To 4 %>
	<% if F_ReturnTime(iCnt) = "1" then %>
<!-- 2009/03/10 R.Shibuta Add-E -->
		<TR>
			<TD width="5%">　</TD>
			<TD width="95%" colspan="11">(9)空コン返却状況</TD>
		</TR>
	<% ItemsToSend = ItemsToSend + 1 %>
<!-- 2009/03/10 R.Shibuta Add-S -->
	<% Exit For %>
	<% end if %>
<% Next %>
<!-- 2009/03/10 R.Shibuta Add-E -->

<% if ItemsToSend = 0 then %>
	<TR>
		<TD colspan="3">　　対象項目設定なし</TD>
	</TR>
<% end if %>
<!-- 2009/07/15 Add-S Fujiyama -->
	<TR><TD><BR></TD></TR>
<!-- 2009/07/15 Add-S Fujiyama -->
<!--</TABLE>
<TABLE border=1 cellPadding=5 cellSpacing=0 width="100%">-->
	<TR>
		<TD colspan="12"><B>配信先</B></TD>
	</TR>
<% if Email1 <> "" then %>
	<TR>
		<TD width="5%">　</TD>
		<TD width="50%" colspan="2">1.　<%=Email1%></TD>
<!-- 2009/03/10 R.Shibuta Add-S -->		
		<% iChk=0 %>
		<% if F_ArrivalTime(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(1)</TD>
		<% end if %>
		<% if F_InTime(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(2)</TD>
		<% end if %>
		<% if F_List(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(3)</TD>
		<% end if %>
		<% if F_DOStatus(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(4)</TD>
		<% end if %>
		<% if F_DelPermit(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(5)</TD>
		<% end if %>
		<% if F_DemurrageFreeTime(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(6)</TD>
		<% end if %>
		<% if F_CYDelTime(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(7)</TD>
		<% end if %>
		<% if F_DetentionFreeTime(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(8)</TD>
		<% end if %>
		<% if F_ReturnTime(0) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(9)</TD>
		<% end if %>
		<% if iChk < 9 then %>
			<TD width=<% =Chr(34) & (9-iChk)*5 & "%" & Chr(34) %> colspan=<% =Chr(34) & (9-iChk) & Chr(34) %>></TD>
		<% end if %>
<!-- 2009/03/10 R.Shibuta Add-E -->
	</TR>
<% end if %>
<% if Email2 <> "" then %>
	<TR>
		<TD width="5%">　</TD>
		<TD width="50%" colspan="2">2.　<%=Email2%></TD>
		
<!-- 2009/03/10 R.Shibuta Add-S -->
		<% iChk=0 %>
		<% if F_ArrivalTime(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(1)</TD>
		<% end if %>
		<% if F_InTime(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(2)</TD>
		<% end if %>
		<% if F_List(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(3)</TD>
		<% end if %>
		<% if F_DOStatus(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(4)</TD>
		<% end if %>
		<% if F_DelPermit(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(5)</TD>
		<% end if %>
		<% if F_DemurrageFreeTime(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(6)</TD>
		<% end if %>
		<% if F_CYDelTime(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(7)</TD>
		<% end if %>
		<% if F_DetentionFreeTime(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(8)</TD>
		<% end if %>
		<% if F_ReturnTime(1) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(9)</TD>
		<% end if %>
		<% if iChk < 9 then %>
			<TD width=<% =Chr(34) & (9-iChk)*5 & "%" & Chr(34) %> colspan=<% =Chr(34) & (9-iChk) & Chr(34) %>></TD>
		<% end if %>
<!-- 2009/03/10 R.Shibuta Add-E -->
	</TR>
<% end if %>
<% if Email3 <> "" then %>
	<TR>
		<TD width="5%">　</TD>
		<TD width="50%" colspan="2">3.　<%=Email3%></TD>
		
<!-- 2009/03/10 R.Shibuta Add-S -->
		<% iChk=0 %>
		<% if F_ArrivalTime(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(1)</TD>
		<% end if %>
		<% if F_InTime(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(2)</TD>
		<% end if %>
		<% if F_List(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(3)</TD>
		<% end if %>
		<% if F_DOStatus(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(4)</TD>
		<% end if %>
		<% if F_DelPermit(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(5)</TD>
		<% end if %>
		<% if F_DemurrageFreeTime(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(6)</TD>
		<% end if %>
		<% if F_CYDelTime(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(7)</TD>
		<% end if %>
		<% if F_DetentionFreeTime(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(8)</TD>
		<% end if %>
		<% if F_ReturnTime(2) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(9)</TD>
		<% end if %>
		<% if iChk < 9 then %>
			<TD width=<% =Chr(34) & (9-iChk)*5 & "%" & Chr(34) %> colspan=<% =Chr(34) & (9-iChk) & Chr(34) %>></TD>
		<% end if %>
<!-- 2009/03/10 R.Shibuta Add-E -->
	</TR>
<% end if %>
<% if Email4 <> "" then %>
	<TR>
		<TD width="5%">　</TD>
		<TD width="50%" colspan="2">4.　<%=Email4%></TD>

<!-- 2009/03/10 R.Shibuta Add-S -->
		<% iChk=0 %>
		<% if F_ArrivalTime(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(1)</TD>
		<% end if %>
		<% if F_InTime(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(2)</TD>
		<% end if %>
		<% if F_List(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(3)</TD>
		<% end if %>
		<% if F_DOStatus(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(4)</TD>
		<% end if %>
		<% if F_DelPermit(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(5)</TD>
		<% end if %>
		<% if F_DemurrageFreeTime(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(6)</TD>
		<% end if %>
		<% if F_CYDelTime(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(7)</TD>
		<% end if %>
		<% if F_DetentionFreeTime(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(8)</TD>
		<% end if %>
		<% if F_ReturnTime(3) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(9)</TD>
		<% end if %>
		<% if iChk < 9 then %>
			<TD width=<% =Chr(34) & (9-iChk)*5 & "%" & Chr(34) %> colspan=<% =Chr(34) & (9-iChk) & Chr(34) %>></TD>
		<% end if %>
<!-- 2009/03/10 R.Shibuta Add-E -->
	</TR>
<% end if %>
<% if Email5 <> "" then %>
	<TR>
		<TD width="5%">　</TD>
		<TD width="50%" colspan="2">5.　<%=Email5%></TD>
		
<!-- 2009/03/10 R.Shibuta Add-S -->
		<% iChk=0 %>
		<% if F_ArrivalTime(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(1)</TD>
		<% end if %>
		<% if F_InTime(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(2)</TD>
		<% end if %>
		<% if F_List(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(3)</TD>
		<% end if %>
		<% if F_DOStatus(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(4)</TD>
		<% end if %>
		<% if F_DelPermit(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(5)</TD>
		<% end if %>
		<% if F_DemurrageFreeTime(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(6)</TD>
		<% end if %>
		<% if F_CYDelTime(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(7)</TD>
		<% end if %>
		<% if F_DetentionFreeTime(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(8)</TD>
		<% end if %>
		<% if F_ReturnTime(4) = "1" then %>
			<% iChk = iChk + 1 %>
			<TD width="5%" colspan="1">(9)</TD>
		<% end if %>
		<% if iChk < 9 then %>
			<TD width=<% =Chr(34) & (9-iChk)*5 & "%" & Chr(34) %> colspan=<% =Chr(34) & (9-iChk) & Chr(34) %>></TD>
		<% end if %>
<!-- 2009/03/10 R.Shibuta Add-E -->
	</TR>
<% end if %>
	<TR>
		<TD width="100%" align="center" colspan="12">
<!-- 2009/03/10 R.Shibuta Upd-S -->
<!--		<INPUT type="hidden" name="F_ArrivalTime" value="<%'=F_ArrivalTime%>"> -->
<!--		<INPUT type="hidden" name="F_InTime" value="<%'=F_InTime%>"> -->
<!--		<INPUT type="hidden" name="F_List" value="<%'=F_List%>"> -->
<!--		<INPUT type="hidden" name="F_DOStatus" value="<%'=F_DOStatus%>"> -->
<!--		<INPUT type="hidden" name="F_DelPermit" value="<%'=F_DelPermit%>"> -->
<!--		<INPUT type="hidden" name="F_DemurrageFreeTime" value="<%'=F_DemurrageFreeTime%>"> -->
<!--		<INPUT type="hidden" name="F_CYDelTime" value="<%'=F_CYDelTime%>"> -->
<!--		<INPUT type="hidden" name="F_DetentionFreeTime" value="<%'=F_DetentionFreeTime%>"> -->
<!--		<INPUT type="hidden" name="F_ReturnTime" value="<%'=F_ReturnTime%>"> -->

			<INPUT type="hidden" name="F_ArrivalTime1" value="<%=F_ArrivalTime(0)%>">
			<INPUT type="hidden" name="F_ArrivalTime2" value="<%=F_ArrivalTime(1)%>">
			<INPUT type="hidden" name="F_ArrivalTime3" value="<%=F_ArrivalTime(2)%>">
			<INPUT type="hidden" name="F_ArrivalTime4" value="<%=F_ArrivalTime(3)%>">
			<INPUT type="hidden" name="F_ArrivalTime5" value="<%=F_ArrivalTime(4)%>">

			<INPUT type="hidden" name="F_InTime1" value="<%=F_InTime(0)%>">
			<INPUT type="hidden" name="F_InTime2" value="<%=F_InTime(1)%>">
			<INPUT type="hidden" name="F_InTime3" value="<%=F_InTime(2)%>">
			<INPUT type="hidden" name="F_InTime4" value="<%=F_InTime(3)%>">
			<INPUT type="hidden" name="F_InTime5" value="<%=F_InTime(4)%>">
			
			<INPUT type="hidden" name="F_List1" value="<%=F_List(0)%>">
			<INPUT type="hidden" name="F_List2" value="<%=F_List(1)%>">
			<INPUT type="hidden" name="F_List3" value="<%=F_List(2)%>">
			<INPUT type="hidden" name="F_List4" value="<%=F_List(3)%>">
			<INPUT type="hidden" name="F_List5" value="<%=F_List(4)%>">
			
			<INPUT type="hidden" name="F_DOStatus1" value="<%=F_DOStatus(0)%>">
			<INPUT type="hidden" name="F_DOStatus2" value="<%=F_DOStatus(1)%>">
			<INPUT type="hidden" name="F_DOStatus3" value="<%=F_DOStatus(2)%>">
			<INPUT type="hidden" name="F_DOStatus4" value="<%=F_DOStatus(3)%>">
			<INPUT type="hidden" name="F_DOStatus5" value="<%=F_DOStatus(4)%>">
						
			<INPUT type="hidden" name="F_DelPermit1" value="<%=F_DelPermit(0)%>">
			<INPUT type="hidden" name="F_DelPermit2" value="<%=F_DelPermit(1)%>">
			<INPUT type="hidden" name="F_DelPermit3" value="<%=F_DelPermit(2)%>">
			<INPUT type="hidden" name="F_DelPermit4" value="<%=F_DelPermit(3)%>">
			<INPUT type="hidden" name="F_DelPermit5" value="<%=F_DelPermit(4)%>">
			
			<INPUT type="hidden" name="F_DemurrageFreeTime1" value="<%=F_DemurrageFreeTime(0)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime2" value="<%=F_DemurrageFreeTime(1)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime3" value="<%=F_DemurrageFreeTime(2)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime4" value="<%=F_DemurrageFreeTime(3)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime5" value="<%=F_DemurrageFreeTime(4)%>">
			
			<INPUT type="hidden" name="DaysToDMFT" value="<%=DaysToDMFT%>">

			<INPUT type="hidden" name="F_CYDelTime1" value="<%=F_CYDelTime(0)%>">
			<INPUT type="hidden" name="F_CYDelTime2" value="<%=F_CYDelTime(1)%>">
			<INPUT type="hidden" name="F_CYDelTime3" value="<%=F_CYDelTime(2)%>">
			<INPUT type="hidden" name="F_CYDelTime4" value="<%=F_CYDelTime(3)%>">
			<INPUT type="hidden" name="F_CYDelTime5" value="<%=F_CYDelTime(4)%>">
			
			<INPUT type="hidden" name="F_DetentionFreeTime1" value="<%=F_DetentionFreeTime(0)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime2" value="<%=F_DetentionFreeTime(1)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime3" value="<%=F_DetentionFreeTime(2)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime4" value="<%=F_DetentionFreeTime(3)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime5" value="<%=F_DetentionFreeTime(4)%>">

			<INPUT type="hidden" name="DaysToDTFT" value="<%=DaysToDTFT%>">
			
			<INPUT type="hidden" name="F_ReturnTime1" value="<%=F_ReturnTime(0)%>">
			<INPUT type="hidden" name="F_ReturnTime2" value="<%=F_ReturnTime(1)%>">
			<INPUT type="hidden" name="F_ReturnTime3" value="<%=F_ReturnTime(2)%>">
			<INPUT type="hidden" name="F_ReturnTime4" value="<%=F_ReturnTime(3)%>">
			<INPUT type="hidden" name="F_ReturnTime5" value="<%=F_ReturnTime(4)%>">
<!-- 2009/03/10 R.Shibuta Upd-E -->
			<INPUT type="hidden" name="Email1" value="<%=Email1%>">
			<INPUT type="hidden" name="Email2" value="<%=Email2%>">
			<INPUT type="hidden" name="Email3" value="<%=Email3%>">
			<INPUT type="hidden" name="Email4" value="<%=Email4%>">
			<INPUT type="hidden" name="Email5" value="<%=Email5%>">
			<INPUT type="submit" value="ＯＫ" onClick="return GoEntry()">
			<INPUT type="submit" value="戻る" onClick="return GoBack()">
		</TD>
	</TR>
</TABLE>
</FORM>
<% else %>
<FORM name="sst301" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
	<TR>
		<TD colspan="3">
			<B>輸入ステータス配信依頼（設定）項目確認</B>
		</TD>
	</TR>
	<TR><TD colspan="3">　</TD></TR>
	<TR>
		<TD colspan="3">
			　　何も指定されていません。よろしければ「ＯＫ」ボタンをクリックしてください。
		</TD>
	</TR>
	<TR><TD colspan="3">　</TD></TR>
	<TR>
		<TD width="100%" align="center" colspan="3">
		
<!-- 2009/03/10 R.Shibuta Upd-S -->
<!--		<INPUT type="hidden" name="F_ArrivalTime1" value="<%'=F_ArrivalTime%>"> -->
<!--		<INPUT type="hidden" name="F_InTime1" value="<%'=F_InTime%>"> -->
<!--		<INPUT type="hidden" name="F_List1" value="<%'=F_List%>"> -->
<!--		<INPUT type="hidden" name="F_DOStatus1" value="<%'=F_DOStatus%>"> -->
<!--		<INPUT type="hidden" name="F_DelPermit1" value="<%'=F_DelPermit%>"> -->
<!--		<INPUT type="hidden" name="F_DemurrageFreeTime1" value="<%'=F_DemurrageFreeTime%>"> -->
<!--		<INPUT type="hidden" name="F_CYDelTime1" value="<%'=F_CYDelTime%>"> -->
<!--		<INPUT type="hidden" name="F_DetentionFreeTime1" value="<%'=F_DetentionFreeTime%>"> -->
<!--		<INPUT type="hidden" name="F_ReturnTime1" value="<%'=F_ReturnTime%>"> -->

			<INPUT type="hidden" name="F_ArrivalTime1" value="<%=F_ArrivalTime(0)%>">
			<INPUT type="hidden" name="F_ArrivalTime2" value="<%=F_ArrivalTime(1)%>">
			<INPUT type="hidden" name="F_ArrivalTime3" value="<%=F_ArrivalTime(2)%>">
			<INPUT type="hidden" name="F_ArrivalTime4" value="<%=F_ArrivalTime(3)%>">
			<INPUT type="hidden" name="F_ArrivalTime5" value="<%=F_ArrivalTime(4)%>">

			<INPUT type="hidden" name="F_InTime1" value="<%=F_InTime(0)%>">
			<INPUT type="hidden" name="F_InTime2" value="<%=F_InTime(1)%>">
			<INPUT type="hidden" name="F_InTime3" value="<%=F_InTime(2)%>">
			<INPUT type="hidden" name="F_InTime4" value="<%=F_InTime(3)%>">
			<INPUT type="hidden" name="F_InTime5" value="<%=F_InTime(4)%>">
			
			<INPUT type="hidden" name="F_List1" value="<%=F_List(0)%>">
			<INPUT type="hidden" name="F_List2" value="<%=F_List(1)%>">
			<INPUT type="hidden" name="F_List3" value="<%=F_List(2)%>">
			<INPUT type="hidden" name="F_List4" value="<%=F_List(3)%>">
			<INPUT type="hidden" name="F_List5" value="<%=F_List(4)%>">
			
			<INPUT type="hidden" name="F_DOStatus1" value="<%=F_DOStatus(0)%>">
			<INPUT type="hidden" name="F_DOStatus2" value="<%=F_DOStatus(1)%>">
			<INPUT type="hidden" name="F_DOStatus3" value="<%=F_DOStatus(2)%>">
			<INPUT type="hidden" name="F_DOStatus4" value="<%=F_DOStatus(3)%>">
			<INPUT type="hidden" name="F_DOStatus5" value="<%=F_DOStatus(4)%>">
						
			<INPUT type="hidden" name="F_DelPermit1" value="<%=F_DelPermit(0)%>">
			<INPUT type="hidden" name="F_DelPermit2" value="<%=F_DelPermit(1)%>">
			<INPUT type="hidden" name="F_DelPermit3" value="<%=F_DelPermit(2)%>">
			<INPUT type="hidden" name="F_DelPermit4" value="<%=F_DelPermit(3)%>">
			<INPUT type="hidden" name="F_DelPermit5" value="<%=F_DelPermit(4)%>">
			
			<INPUT type="hidden" name="F_DemurrageFreeTime1" value="<%=F_DemurrageFreeTime(0)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime2" value="<%=F_DemurrageFreeTime(1)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime3" value="<%=F_DemurrageFreeTime(2)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime4" value="<%=F_DemurrageFreeTime(3)%>">
			<INPUT type="hidden" name="F_DemurrageFreeTime5" value="<%=F_DemurrageFreeTime(4)%>">
			
			<INPUT type="hidden" name="DaysToDMFT" value="<%=DaysToDMFT%>">

			<INPUT type="hidden" name="F_CYDelTime1" value="<%=F_CYDelTime(0)%>">
			<INPUT type="hidden" name="F_CYDelTime2" value="<%=F_CYDelTime(1)%>">
			<INPUT type="hidden" name="F_CYDelTime3" value="<%=F_CYDelTime(2)%>">
			<INPUT type="hidden" name="F_CYDelTime4" value="<%=F_CYDelTime(3)%>">
			<INPUT type="hidden" name="F_CYDelTime5" value="<%=F_CYDelTime(4)%>">
			
			<INPUT type="hidden" name="F_DetentionFreeTime1" value="<%=F_DetentionFreeTime(0)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime2" value="<%=F_DetentionFreeTime(1)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime3" value="<%=F_DetentionFreeTime(2)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime4" value="<%=F_DetentionFreeTime(3)%>">
			<INPUT type="hidden" name="F_DetentionFreeTime5" value="<%=F_DetentionFreeTime(4)%>">

			<INPUT type="hidden" name="DaysToDTFT" value="<%=DaysToDTFT%>">
			
			<INPUT type="hidden" name="F_ReturnTime1" value="<%=F_ReturnTime(0)%>">
			<INPUT type="hidden" name="F_ReturnTime2" value="<%=F_ReturnTime(1)%>">
			<INPUT type="hidden" name="F_ReturnTime3" value="<%=F_ReturnTime(2)%>">
			<INPUT type="hidden" name="F_ReturnTime4" value="<%=F_ReturnTime(3)%>">
			<INPUT type="hidden" name="F_ReturnTime5" value="<%=F_ReturnTime(4)%>">
<!-- 2009/03/10 R.Shibuta Upd-E -->

			<INPUT type="hidden" name="Email1" value="<%=Email1%>">
			<INPUT type="hidden" name="Email2" value="<%=Email2%>">
			<INPUT type="hidden" name="Email3" value="<%=Email3%>">
			<INPUT type="hidden" name="Email4" value="<%=Email4%>">
			<INPUT type="hidden" name="Email5" value="<%=Email5%>">
			<INPUT type="submit" value="ＯＫ" onClick="return GoEntry()">
			<INPUT type="submit" value="戻る" onClick="return GoBack()">
		</TD>		
	</TR>
</TABLE>
</FORM>
<% end if %>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
