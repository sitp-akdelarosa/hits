<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst300.asp				_/
'_/	Function	:ステータス配信対象項目設定画面			_/
'_/	Date			:2003/12/26				_/
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
	CheckLoginH2

'データ取得

'2009/03/10 R.Shibuta Upd-S
'	Dim F_ArrivalTime, F_InTime, F_List, F_DOStatus, F_DelPermit
'	Dim F_DemurrageFreeTime, F_CYDelTime, F_DetentionFreeTime, F_ReturnTime
	Dim F_ArrivalTime(4), F_InTime(4), F_List(4), F_DOStatus(4), F_DelPermit(4)
	Dim F_DemurrageFreeTime(4), F_CYDelTime(4), F_DetentionFreeTime(4), F_ReturnTime(4)
'2009/03/10 R.Shibuta Upd-E
	Dim iCnt
	
	Dim DaysToDMFT, DaysToDTFT
	Dim Email1, Email2, Email3, Email4, Email5
	Dim USER
	Dim chkDays


	'入力内容の確認画面からの戻りでない場合、
	'すなわち「設定」メニューから来た場合にのみＤＢから値を取得する。
	if Session.Contents("sst301") <> "true" then

		USER = Session.Contents("userid")

	'DB接続
		Dim ObjConn, ObjRS, StrSQL
		ConnDBH ObjConn, ObjRS

		StrSQL = "SELECT * from TargetItems where UserCode='"& USER &"'"
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB切断
			jumpErrorP "1","c103","01","ステータス配信対象項目設定","101","SQL:<BR>"&strSQL
		end if

		if ObjRS.eof then
		'2009/03/10 R.Shibuta Upd-S
'		if ObjRS.eof then
'			F_ArrivalTime = ""
'			F_InTime = ""
'			F_List = ""
'			F_DOStatus = ""
'			F_DelPermit = ""
'			F_DemurrageFreeTime = ""
'			DaysToDMFT = ""
'			F_CYDelTime = ""
'			F_DetentionFreeTime = ""
'			DaysToDTFT = ""
'			F_ReturnTime = ""
'			Email1 = ""
'			Email2 = ""
'			Email3 = ""
'			Email4 = ""
'			Email5 = ""
'		else
'			F_ArrivalTime = ObjRS("FlagArrivalTime")
'			F_InTime = ObjRS("FlagInTime")
'			F_List = ObjRS("FlagList")
'			F_DOStatus = ObjRS("FlagDOStatus")
'			F_DelPermit = ObjRS("FlagDelPermit")
'			F_DemurrageFreeTime = ObjRS("FlagDemurrageFreeTime")
'			DaysToDMFT = Trim(ObjRS("DaysToDemurrageFreeTime"))
'			F_CYDelTime = ObjRS("FlagCYDelTime")
'			F_DetentionFreeTime = ObjRS("FlagDetentionFreeTime")
'			DaysToDTFT = Trim(ObjRS("DaysToDetentionFreeTime"))
'			F_ReturnTime = ObjRS("FlagReturnTime")
'			Email1 = Trim(ObjRS("Email1"))
'			Email2 = Trim(ObjRS("Email2"))
'			Email3 = Trim(ObjRS("Email3"))
'			Email4 = Trim(ObjRS("Email4"))
'			Email5 = Trim(ObjRS("Email5"))
'		end if
			For icnt = 0 TO 4
				F_ArrivalTime(icnt) = ""
				F_InTime(icnt) = ""
				F_List(icnt) = ""
				F_DOStatus(icnt) = ""
				F_DelPermit(icnt) = ""
				F_DemurrageFreeTime(icnt) = ""
				F_CYDelTime(icnt) = ""
				F_DetentionFreeTime(icnt) = ""
				F_ReturnTime(icnt) = ""
			Next
				DaysToDMFT = ""
				DaysToDTFT = ""
				Email1 = ""
				Email2 = ""
				Email3 = ""
				Email4 = ""
				Email5 = ""
		else
		
			F_ArrivalTime(0) = ObjRS("FlagArrivalTime")
			F_ArrivalTime(1) = ObjRS("FlagArrivalTime2")
			F_ArrivalTime(2) = ObjRS("FlagArrivalTime3")
			F_ArrivalTime(3) = ObjRS("FlagArrivalTime4")
			F_ArrivalTime(4) = ObjRS("FlagArrivalTime5")
		
			F_InTime(0) = ObjRS("FlagInTime")
			F_InTime(1) = ObjRS("FlagInTime2")
			F_InTime(2) = ObjRS("FlagInTime3")
			F_InTime(3) = ObjRS("FlagInTime4")
			F_InTime(4) = ObjRS("FlagInTime5")
		
			F_List(0) = ObjRS("FlagList")
			F_List(1) = ObjRS("FlagList2")
			F_List(2) = ObjRS("FlagList3")
			F_List(3) = ObjRS("FlagList4")
			F_List(4) = ObjRS("FlagList5")
		
			F_DOStatus(0) = ObjRS("FlagDOStatus")
			F_DOStatus(1) = ObjRS("FlagDOStatus2")
			F_DOStatus(2) = ObjRS("FlagDOStatus3")
			F_DOStatus(3) = ObjRS("FlagDOStatus4")
			F_DOStatus(4) = ObjRS("FlagDOStatus5")
		
			F_DelPermit(0) = ObjRS("FlagDelPermit")
			F_DelPermit(1) = ObjRS("FlagDelPermit2")
			F_DelPermit(2) = ObjRS("FlagDelPermit3")
			F_DelPermit(3) = ObjRS("FlagDelPermit4")
			F_DelPermit(4) = ObjRS("FlagDelPermit5")
		
			F_DemurrageFreeTime(0) = ObjRS("FlagDemurrageFreeTime")
			F_DemurrageFreeTime(1) = ObjRS("FlagDemurrageFreeTime2")
			F_DemurrageFreeTime(2) = ObjRS("FlagDemurrageFreeTime3")
			F_DemurrageFreeTime(3) = ObjRS("FlagDemurrageFreeTime4")
			F_DemurrageFreeTime(4) = ObjRS("FlagDemurrageFreeTime5")
		
			F_CYDelTime(0) = ObjRS("FlagCYDelTime")
			F_CYDelTime(1) = ObjRS("FlagCYDelTime2")
			F_CYDelTime(2) = ObjRS("FlagCYDelTime3")
			F_CYDelTime(3) = ObjRS("FlagCYDelTime4")
			F_CYDelTime(4) = ObjRS("FlagCYDelTime5")
		
			F_DetentionFreeTime(0) = ObjRS("FlagDetentionFreeTime")
			F_DetentionFreeTime(1) = ObjRS("FlagDetentionFreeTime2")
			F_DetentionFreeTime(2) = ObjRS("FlagDetentionFreeTime3")
			F_DetentionFreeTime(3) = ObjRS("FlagDetentionFreeTime4")
			F_DetentionFreeTime(4) = ObjRS("FlagDetentionFreeTime5")
		
			F_ReturnTime(0) = ObjRS("FlagReturnTime")
			F_ReturnTime(1) = ObjRS("FlagReturnTime2")
			F_ReturnTime(2) = ObjRS("FlagReturnTime3")
			F_ReturnTime(3) = ObjRS("FlagReturnTime4")
			F_ReturnTime(4) = ObjRS("FlagReturnTime5")
			
			DaysToDMFT = Trim(ObjRS("DaysToDemurrageFreeTime"))
			DaysToDTFT = Trim(ObjRS("DaysToDetentionFreeTime"))
			Email1 = Trim(ObjRS("Email1"))
			Email2 = Trim(ObjRS("Email2"))
			Email3 = Trim(ObjRS("Email3"))
			Email4 = Trim(ObjRS("Email4"))
			Email5 = Trim(ObjRS("Email5"))
		end if
		'2009/03/10 R.Shibuta Upd-E
		ObjRS.close

	'DB接続解除
		DisConnDBH ObjConn, ObjRS
	'エラートラップ解除
		on error goto 0

	'ログ出力
	WriteLogH "c103", "ステータス配信対象設定","01",""

	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>IMPORT DELIVERY REQUEST (SETUP)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(900,900);

window.focus();

function GoStop(){
<%	Session.Contents("sst301") = "false" %>
	window.close();
}

function GoEntry(){

	f=document.sst300;
	//何も選択されていない場合
//	if((f.ArrivalTime.checked==false) &&
//		 (f.InTime.checked==false) &&
//		 (f.List.checked==false) &&
//		 (f.DOStatus.checked==false) &&
//		 (f.DelPermit.checked==false) &&
//		 (f.DemurrageFreeTime.checked==false) &&
//		 (f.DaysToDemurrageFreeTime.value=="") &&
//		 (f.CYDelTime.checked==false) &&
//		 (f.DetentionFreeTime.checked==false) &&
//		 (f.DaysToDetentionFreeTime.value=="") &&
//		 (f.ReturnTime.checked==false)){
//			alert("何も選択されていません。");
//			f.ArrivalTime.focus();
//			return false;
//	}
//		2009/03/10 R.Shibuta Upd-S
		//デマレージフリータイムがチェックされているが
		if(f.DemurrageFreeTime1.checked==true || 
		f.DemurrageFreeTime2.checked==true || 
		f.DemurrageFreeTime3.checked==true || 
		f.DemurrageFreeTime4.checked==true || 
		f.DemurrageFreeTime5.checked==true){
//		2009/03/10 R.Shibuta Upd-E
			//日数に入力がない場合
			if(f.DaysToDemurrageFreeTime.value==""){
				alert("デマレージフリータイムのチェック日数を入力してください。");
				f.DaysToDemurrageFreeTime.focus();
				return false;
			}
			//日数に数字以外が入力されている場合
			if(isNaN(f.DaysToDemurrageFreeTime.value)){
				alert("デマレージフリータイムのチェック日数を数字で入力してください。");
				f.DaysToDemurrageFreeTime.focus();
				return false;
			}
			//日数に半角スペースが入力されている場合
			if(CheckSpace(f.DaysToDemurrageFreeTime.value)){
				alert("デマレージフリータイムのチェック日数を数字で入力してください。");
				f.DaysToDemurrageFreeTime.focus();
				return false;
			}
			//日数に0が入力されている場合
			if(f.DaysToDemurrageFreeTime.value == "0"){
				alert("デマレージフリータイムのチェック日数には１以上を入力してください。");
				f.DaysToDemurrageFreeTime.focus();
				return false;
			}
		}

		//デマレージフリータイムの日数だけ入力され、チェックボックスがチェックされていない場合
//		2009/03/10 R.Shibuta Upd-S
		if((f.DemurrageFreeTime1.checked==false &&
		    f.DemurrageFreeTime2.checked==false &&
		    f.DemurrageFreeTime3.checked==false &&
		    f.DemurrageFreeTime4.checked==false &&
		    f.DemurrageFreeTime5.checked==false) &&
//		2009/03/10 R.Shibuta Upd-E
			 (f.DaysToDemurrageFreeTime.value!="")){
				alert("デマレージフリータイムをチェックするか日数を削除してください。");
				f.DaysToDemurrageFreeTime.focus();
				return false;
		}
				
		//ディテンションフリータイムがチェックされているが
//		2009/03/10 R.Shibuta Upd-S
		if(f.DetentionFreeTime1.checked==true || 
		f.DetentionFreeTime2.checked==true || 
		f.DetentionFreeTime3.checked==true || 
		f.DetentionFreeTime4.checked==true || 
		f.DetentionFreeTime5.checked==true){
//		2009/03/10 R.Shibuta Upd-E
			//日数に入力がない場合
			if(f.DaysToDetentionFreeTime.value==""){
				alert("ディテンションフリータイムのチェック日数を入力してください。");
				f.DaysToDetentionFreeTime.focus();
				return false;
			}
			//日数に数字以外が入力されている場合
			if(isNaN(f.DaysToDetentionFreeTime.value)){
				alert("ディテンションフリータイムのチェック日数を数字で入力してください。");
				f.DaysToDetentionFreeTime.focus();
				return false;
			}
			//日数に半角スペースが入力されている場合
			if(CheckSpace(f.DaysToDetentionFreeTime.value)){
				alert("ディテンションフリータイムのチェック日数を数字で入力してください。");
				f.DaysToDetentionFreeTime.focus();
				return false;
			}
			//日数に0が入力されている場合
			if(f.DaysToDetentionFreeTime.value == "0"){
				alert("ディテンションフリータイムのチェック日数には１以上を入力してください。");
				f.DaysToDetentionFreeTime.focus();
				return false;
			}
		}
		
		//ディテンションフリータイムの日数だけ入力され、チェックボックスがチェックされていない場合
//		2009/03/10 R.Shibuta Upd-S
		if((f.DetentionFreeTime1.checked==false &&
		    f.DetentionFreeTime2.checked==false &&
		    f.DetentionFreeTime3.checked==false &&
		    f.DetentionFreeTime4.checked==false &&
		    f.DetentionFreeTime5.checked==false) &&
//		2009/03/10 R.Shibuta Upd-E
			 (f.DaysToDetentionFreeTime.value!="")){
				alert("ディテンションフリータイムをチェックするか日数を削除してください。");
				f.DaysToDetentionFreeTime.focus();
				return false;
		}
		
	//メールアドレスが何も入力されていない場合
//	if((f.Email1.value=="") &&
//		 (f.Email2.value=="") &&
//		 (f.Email3.value=="") &&
//		 (f.Email4.value=="") &&
//		 (f.Email5.value=="")){
//			alert("メールアドレスを入力してください。");
//			f.Email1.focus();
//			return false;
//	}
	//メールアドレスの内容チェック
	if(f.Email1.value!=""){
		if(gfisMailAddr(f.Email1.value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			f.Email1.focus();
			return false;
		}
		if(f.Email1.value==f.Email2.value || f.Email1.value==f.Email3.value ||
			 f.Email1.value==f.Email4.value || f.Email1.value==f.Email5.value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				f.Email1.focus();
				return false;
			}
		}
	}
	if(f.Email2.value!=""){
		if(gfisMailAddr(f.Email2.value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			f.Email2.focus();
			return false;
		}
		if(f.Email2.value==f.Email3.value || f.Email2.value==f.Email4.value ||
			 f.Email2.value==f.Email5.value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				f.Email2.focus();
				return false;
			}
		}
	}
	if(f.Email3.value!=""){
		if(gfisMailAddr(f.Email3.value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			f.Email3.focus();
			return false;
		}
		if(f.Email3.value==f.Email4.value || f.Email3.value==f.Email5.value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				f.Email3.focus();
				return false;
			}
		}
	}
	if(f.Email4.value!=""){
		if(gfisMailAddr(f.Email4.value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			f.Email4.focus();
			return false;
		}
		if(f.Email4.value==f.Email5.value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				f.Email4.focus();
				return false;
			}
		}
	}
	if((f.Email5.value!="") && (gfisMailAddr(f.Email5.value)==false)){
		alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
		f.Email5.focus();
		return false;
	}
//		2009/03/10 R.Shibuta Upd-S
		if(f.ArrivalTime1.checked==true){
			f.F_ArrivalTime1.value="1"
		}else{
			f.F_ArrivalTime1.value="0"
		}
		if(f.ArrivalTime2.checked==true){
			f.F_ArrivalTime2.value="1"
		}else{
			f.F_ArrivalTime2.value="0"
		}
		if(f.ArrivalTime3.checked==true){
			f.F_ArrivalTime3.value="1"
		}else{
			f.F_ArrivalTime3.value="0"
		}
		if(f.ArrivalTime4.checked==true){
			f.F_ArrivalTime4.value="1"
		}else{
			f.F_ArrivalTime4.value="0"
		}
		if(f.ArrivalTime5.checked==true){
			f.F_ArrivalTime5.value="1"
		}else{
			f.F_ArrivalTime5.value="0"
		}
		
		if(f.InTime1.checked==true){
			f.F_InTime1.value="1"
		}else{
			f.F_InTime1.value="0"
		}
		if(f.InTime2.checked==true){
			f.F_InTime2.value="1"
		}else{
			f.F_InTime2.value="0"
		}
		if(f.InTime3.checked==true){
			f.F_InTime3.value="1"
		}else{
			f.F_InTime3.value="0"
		}
		if(f.InTime4.checked==true){
			f.F_InTime4.value="1"
		}else{
			f.F_InTime4.value="0"
		}
		if(f.InTime5.checked==true){
			f.F_InTime5.value="1"
		}else{
			f.F_InTime5.value="0"
		}		
		if(f.List1.checked==true){
			f.F_List1.value="1"
		}else{
			f.F_List1.value="0"
		}
		if(f.List2.checked==true){
			f.F_List2.value="1"
		}else{
			f.F_List2.value="0"
		}
		if(f.List3.checked==true){
			f.F_List3.value="1"
		}else{
			f.F_List3.value="0"
		}
		if(f.List4.checked==true){
			f.F_List4.value="1"
		}else{
			f.F_List4.value="0"
		}
		if(f.List5.checked==true){
			f.F_List5.value="1"
		}else{
			f.F_List5.value="0"
		}
		
		if(f.DOStatus1.checked==true){
			f.F_DOStatus1.value="1"
		}else{
			f.F_DOStatus1.value="0"
		}
		if(f.DOStatus2.checked==true){
			f.F_DOStatus2.value="1"
		}else{
			f.F_DOStatus2.value="0"
		}
		if(f.DOStatus3.checked==true){
			f.F_DOStatus3.value="1"
		}else{
			f.F_DOStatus3.value="0"
		}
		if(f.DOStatus4.checked==true){
			f.F_DOStatus4.value="1"
		}else{
			f.F_DOStatus4.value="0"
		}
		if(f.DOStatus5.checked==true){
			f.F_DOStatus5.value="1"
		}else{
			f.F_DOStatus5.value="0"
		}		
		if(f.DelPermit1.checked==true){
			f.F_DelPermit1.value="1"
		}else{
			f.F_DelPermit1.value="0"
		}
		if(f.DelPermit2.checked==true){
			f.F_DelPermit2.value="1"
		}else{
			f.F_DelPermit2.value="0"
		}
		if(f.DelPermit3.checked==true){
			f.F_DelPermit3.value="1"
		}else{
			f.F_DelPermit3.value="0"
		}
		if(f.DelPermit4.checked==true){
			f.F_DelPermit4.value="1"
		}else{
			f.F_DelPermit4.value="0"
		}
		if(f.DelPermit5.checked==true){
			f.F_DelPermit5.value="1"
		}else{
			f.F_DelPermit5.value="0"
		}
		
		if(f.DemurrageFreeTime1.checked==true){
			f.F_DemurrageFreeTime1.value="1"
		}else{
			f.F_DemurrageFreeTime1.value="0"
		}
		if(f.DemurrageFreeTime2.checked==true){
			f.F_DemurrageFreeTime2.value="1"
		}else{
			f.F_DemurrageFreeTime2.value="0"
		}
		if(f.DemurrageFreeTime3.checked==true){
			f.F_DemurrageFreeTime3.value="1"
		}else{
			f.F_DemurrageFreeTime3.value="0"
		}
		if(f.DemurrageFreeTime4.checked==true){
			f.F_DemurrageFreeTime4.value="1"
		}else{
			f.F_DemurrageFreeTime4.value="0"
		}
		if(f.DemurrageFreeTime5.checked==true){
			f.F_DemurrageFreeTime5.value="1"
		}else{
			f.F_DemurrageFreeTime5.value="0"
		}
		
		if(f.CYDelTime1.checked==true){
			f.F_CYDelTime1.value="1"
		}else{
			f.F_CYDelTime1.value="0"
		}
		if(f.CYDelTime2.checked==true){
			f.F_CYDelTime2.value="1"
		}else{
			f.F_CYDelTime2.value="0"
		}
		if(f.CYDelTime3.checked==true){
			f.F_CYDelTime3.value="1"
		}else{
			f.F_CYDelTime3.value="0"
		}
		if(f.CYDelTime4.checked==true){
			f.F_CYDelTime4.value="1"
		}else{
			f.F_CYDelTime4.value="0"
		}
		if(f.CYDelTime5.checked==true){
			f.F_CYDelTime5.value="1"
		}else{
			f.F_CYDelTime5.value="0"
		}
		
		if(f.DetentionFreeTime1.checked==true){
			f.F_DetentionFreeTime1.value="1"
		}else{
			f.F_DetentionFreeTime1.value="0"
		}
		if(f.DetentionFreeTime2.checked==true){
			f.F_DetentionFreeTime2.value="1"
		}else{
			f.F_DetentionFreeTime2.value="0"
		}
		if(f.DetentionFreeTime3.checked==true){
			f.F_DetentionFreeTime3.value="1"
		}else{
			f.F_DetentionFreeTime3.value="0"
		}
		if(f.DetentionFreeTime4.checked==true){
			f.F_DetentionFreeTime4.value="1"
		}else{
			f.F_DetentionFreeTime4.value="0"
		}
		if(f.DetentionFreeTime5.checked==true){
			f.F_DetentionFreeTime5.value="1"
		}else{
			f.F_DetentionFreeTime5.value="0"
		}

		if(f.ReturnTime1.checked==true){
			f.F_ReturnTime1.value="1"
		}else{
			f.F_ReturnTime1.value="0"
		}
		if(f.ReturnTime2.checked==true){
			f.F_ReturnTime2.value="1"
		}else{
			f.F_ReturnTime2.value="0"
		}
		if(f.ReturnTime3.checked==true){
			f.F_ReturnTime3.value="1"
		}else{
			f.F_ReturnTime3.value="0"
		}
		if(f.ReturnTime4.checked==true){
			f.F_ReturnTime4.value="1"
		}else{
			f.F_ReturnTime4.value="0"
		}
		if(f.ReturnTime5.checked==true){
			f.F_ReturnTime5.value="1"
		}else{
			f.F_ReturnTime5.value="0"
		}
//		2009/03/10 R.Shibuta Upd-E
	f.action="sst301.asp";
	return true;
}

//メールアドレスチェック
function gfisMailAddr(a){
	if(a==""){
		return(true);
	}
	var b=a.replace(/[a-zA-Z0-9_@\.\-]/g,'');
	if(b.length!=0){
		return(false);
	}
	var p1=a.indexOf("@");
	var p2=a.lastIndexOf("@");
	var p3=a.lastIndexOf(".");
	if(0<p1 && p1==p2 && p1<p3 && p3<a.length-1 ){
		return(true);
	}
	return(false);
}
// 半角スペースチェック
function CheckSpace(checkString){
	len = checkString.length;
	for(var i = 0; i < len; i++){
		ch = checkString.substring(i, i+1);
		if(ch == " "){
			continue;
		}else{
			return false;
		}
	}
	return true;
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信対象項目設定画面--------------------------->
<%'データ登録／更新しました画面にて「最新の情報に更新」でSubmitされた場合の対策 %>
<% Session.Contents("ItemsSubmitted")="False"  %>
<FORM name="sst300" method="POST">
<TABLE border="0" cellPadding="5" cellSpacing="0" width="100%">
	<TR>
		<TD colspan="3">
			<B>IMPORT DELIVERY REQUEST （SETUP）</B>
		</TD>
	</TR>
	<TR><TD colspan="3">　</TD></TR>
	<TR>
		<TD colspan="3">
			Information will be sent whenever the status changes as below.
		</TD>
	</TR>
	<TR>
		<TD width="5%">　</TD>
<% if F_ArrivalTime(0) = "1" or Request.Form("F_ArrivalTime(0)") = "1" then %>

		<TD width="95%" colspan="2">(1) Arrival Time
		    <!--
            <BR>
		　　　<FONT class="font80pct">（KACCS側で時間が未入力、あるいは、0:00:00を入力された場合は、<BR></FONT>
		　　　<FONT class="font80pct">　HiTSでは両者の違いを判断できませんので、年月日までの情報を送信します。）</FONT>
		    -->
		</TD>
<% else %>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!--	<TD width="95%" colspan="2"><input type="checkbox" value="0" name="ArrivalTime">(1)入港時間</TD> -->
		<TD width="95%" colspan="2">(1)Arrival Time</TD>
<!--	2009/03/10 R.Shibuta Upd-E -->
<% end if %>

	</TR>
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!--	<%' if F_InTime = "1" or Request.Form("F_InTime") = "1" then %> -->
<!--			<TD width="95%" colspan="2"><input type="checkbox" value="1" name="InTime" checked>(2)ＣＹ搬入時間</TD> -->
<!--	<%' else %> -->
<!--			<TD width="95%" colspan="2"><input type="checkbox" value="0" name="InTime">(2)ＣＹ搬入時間</TD> -->
<!--	<%' end if %> -->
		<TD width="95%" colspan="2">(2) CY In Time</TD>
<!--	2009/03/10 R.Shibuta Upd-E -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!--	<%' if F_List = "1" or Request.Form("F_List") = "1" then %> -->
<!--			<TD width="95%" colspan="2"><input type="checkbox" value="1" name="List" checked>(3)通関許可状況</TD> -->
<!--	<%' else %> -->
<!--			<TD width="95%" colspan="2"><input type="checkbox" value="0" name="List">(3)通関許可状況</TD> -->
<!--	<% 'end if %> -->
<!--	<TD width="95%" colspan="2"><input type="checkbox" value="1" name="List" checked>(3)通関許可状況</TD> -->
		<TD width="95%" colspan="2">(3) Customs Clearance Status</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!-- <% 'if F_DOStatus = "1" or Request.Form("F_DOStatus") = "1" then %> -->
<!--		<TD width="95%" colspan="2"><input type="checkbox" value="1" name="DOStatus" checked>(4)ＤＯクリア状況</TD> -->
<!-- <%' else %> -->
<!--		<TD width="95%" colspan="2"><input type="checkbox" value="0" name="DOStatus">(4)ＤＯクリア状況</TD> -->
<!--<%' end if %> -->
		<TD width="95%" colspan="2">(4) DO Issue Status</TD>
<!--	</TR> -->
<!--	2009/03/10 R.Shibuta Upd-E -->
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!-- <%' if F_DelPermit = "1" or Request.Form("F_DelPermit") = "1" then %> -->
<!--		<TD width="95%" colspan="2"><input type="checkbox" value="1" name="DelPermit" checked>(5)搬出可否</TD> -->
<!-- <%' else %> -->
<!--		<TD width="95%" colspan="2"><input type="checkbox" value="0" name="DelPermit">(5)搬出可否</TD> -->
<!-- <%' end if %> -->
<!--	2009/03/10 R.Shibuta Upd-E -->
		<TD width="95%" colspan="2">(5) Delivery Permission Status</TD>
	</TR>
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!-- <%' if F_DemurrageFreeTime = "1" or Request.Form("F_DemurrageFreeTime") = "1" then %> -->
<!-- 		<TD width="40%"><input type="checkbox" value="1" name="DemurrageFreeTime" checked>(6)デマレージフリータイム</TD> -->
<!-- <%' else %> -->
<!--		<TD width="40%"><input type="checkbox" value="0" name="DemurrageFreeTime">(6)デマレージフリータイム</TD> -->
<!-- <%' end if %> -->
		<TD width="55%">(6) Demurrage Free Time</TD>
<!--	2009/03/10 R.Shibuta Upd-E -->

<% if Request.Form("DaysToDMFT") <> "" then %>
		<TD width="50%"><input type="text" name="DaysToDemurrageFreeTime" value="<%=Request.Form("DaysToDMFT")%>" size="2" maxlength="1">days to go</TD>
<% else %>
		<TD width="50%"><input type="text" name="DaysToDemurrageFreeTime" value="<%=DaysToDMFT%>" size="2" maxlength="1">days to go</TD>
<% end if %>
	</TR>
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!-- <%' if F_CYDelTime = "1" or Request.Form("F_CYDelTime") = "1" then %> -->
<!--		<TD width="95%" colspan="2"><input type="checkbox" value="1" name="CYDelTime" checked>(7)ＣＹ搬出時間</TD> -->
<!-- <%' else %> -->
<!--		<TD width="95%" colspan="2"><input type="checkbox" value="0" name="CYDelTime">(7)ＣＹ搬出時間</TD> -->
<!-- <%' end if %> -->
<!--	2009/03/10 R.Shibuta Upd-E -->
		<TD width="95%" colspan="2">(7) CY Out Time</TD>
	</TR>
	<TR>
		<TD width="5%">　</TD>
<!--	2009/03/10 R.Shibuta Upd-S -->
<!-- <%' if F_DetentionFreeTime = "1" or Request.Form("F_DetentionFreeTime") = "1" then %> -->
<!--		<TD width="40%"><input type="checkbox" value="1" name="DetentionFreeTime" checked>(8)ディテンションフリータイム</TD> -->
<!-- <%' else %> -->
<!--		<TD width="40%"><input type="checkbox" value="0" name="DetentionFreeTime">(8)ディテンションフリータイム</TD> -->
<!-- <%' end if %> -->
<!--	2009/03/10 R.Shibuta Upd-E -->
		<TD width="45%">(8) Detention Free Time </TD>
		
<% if Request.Form("DaysToDTFT") <> "" then %>
		<TD width="50%"><input type="text" name="DaysToDetentionFreeTime" value="<%=Request.Form("DaysToDTFT")%>" size="2" maxlength="1">days to go</TD>
<% else %>
		<TD width="50%"><input type="text" name="DaysToDetentionFreeTime" value="<%=DaysToDTFT%>" size="2" maxlength="1">days to go</TD>
<% end if %>
	</TR>
	<TR>
		<TD width="5%">　</TD>
		<TD width="95%" colspan="2">(9) Empty Container Return Status</TD>
	</TR>
	<TR>
		<TD height="5">　</TD>
	</TR>	
	<TR>
		<TD width="100%" colspan="3"><B>DELIVERY EMAIL ADDRESS</B></TD>
<!-- 2009/07/15 Add-S Fujiyama -->
		<TR></TR>
		<TD width="10%" colspan="1"></TD>
		<TD width="10%" colspan="2">Check the status items from 1 to 9 to receieve the notice for each address.</TD>
<!-- 2009/07/15 Add-E Fujiyama -->
<!--	2009/03/10 R.Shibuta Add-S -->
		<TD width="5%" align="center" colspan="1">(1)</TD>
		<TD width="5%" align="center" colspan="1">(2)</TD>
		<TD width="5%" align="center" colspan="1">(3)</TD>
		<TD width="5%" align="center" colspan="1">(4)</TD>
		<TD width="5%" align="center" colspan="1">(5)</TD>
		<TD width="5%" align="center" colspan="1">(6)</TD>
		<TD width="5%" align="center" colspan="1">(7)</TD>
		<TD width="5%" align="center" colspan="1">(8)</TD>
		<TD width="5%" align="center" colspan="1">(9)</TD>
<!--	2009/03/10 R.Shibuta Add-E -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
		
<% if Request.Form("Email1") <> "" then %>
		<TD width="95%" colspan="2">1　<input type="text" name="Email1" value="<%=Request.Form("Email1")%>" size="70" maxlength="100"></TD>
<% else %>
		<TD width="95%" colspan="2">1　<input type="text" name="Email1" value="<%=Email1%>" size="70" maxlength="100"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-S -->
<! 入港時間 -->
<% if F_ArrivalTime(0) = "1" or Request.Form("F_ArrivalTime1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ArrivalTime1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ArrivalTime1"></TD>
<% end if %>

<! CY搬入時間 -->
<% if F_InTime(0) = "1" or Request.Form("F_InTime1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="InTime1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="InTime1"></TD>
<% end if %>

<! 通関許可状況 -->
<% if F_List(0) = "1" or Request.Form("F_List1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="List1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="List1"></TD>
<% end if %>

<! ＤＯクリア状況 -->
<% if F_DOStatus(0) = "1" or Request.Form("F_DOStatus1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DOStatus1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DOStatus1"></TD>
<% end if %>

<! 搬出可否 -->
<% if F_DelPermit(0) = "1" or Request.Form("F_DelPermit1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DelPermit1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DelPermit1"></TD>
<% end if %>

<! デマレージフリータイム -->
<% if F_DemurrageFreeTime(0) = "1" or Request.Form("F_DemurrageFreeTime1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DemurrageFreeTime1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DemurrageFreeTime1"></TD>
<% end if %>

<! CY搬出時間 -->
<% if F_CYDelTime(0) = "1" or Request.Form("F_CYDelTime1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="CYDelTime1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="CYDelTime1"></TD>
<% end if %>

<! ディテンションフリータイム -->
<% if F_DetentionFreeTime(0) = "1" or Request.Form("F_DetentionFreeTime1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DetentionFreeTime1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DetentionFreeTime1"></TD>
<% end if %>

<! 空コン返却状況 -->
<% if F_ReturnTime(0) = "1" or Request.Form("F_ReturnTime1") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ReturnTime1" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ReturnTime1"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-E -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
<% if Request.Form("Email2") <> "" then %>
		<TD width="95%" colspan="2">2　<input type="text" name="Email2" value="<%=Request.Form("Email2")%>" size="70" maxlength="100"></TD>
<% else %>
		<TD width="95%" colspan="2">2　<input type="text" name="Email2" value="<%=Email2%>" size="70" maxlength="100"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-S -->
<! 入港時間 -->
<% if F_ArrivalTime(1) = "1" or Request.Form("F_ArrivalTime2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ArrivalTime2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ArrivalTime2"></TD>
<% end if %>

<! CY搬入時間 -->
<% if F_InTime(1) = "1" or Request.Form("F_InTime2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="InTime2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="InTime2"></TD>
<% end if %>

<! 通関許可状況 -->
<% if F_List(1) = "1" or Request.Form("F_List2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="List2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="List2"></TD>
<% end if %>

<! ＤＯクリア状況 -->
<% if F_DOStatus(1) = "1" or Request.Form("F_DOStatus2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DOStatus2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DOStatus2"></TD>
<% end if %>

<! 搬出可否 -->
<% if F_DelPermit(1) = "1" or Request.Form("F_DelPermit2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DelPermit2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DelPermit2"></TD>
<% end if %>

<! デマレージフリータイム -->
<% if F_DemurrageFreeTime(1) = "1" or Request.Form("F_DemurrageFreeTime2") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DemurrageFreeTime2" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DemurrageFreeTime2"></TD>
<% end if %>

<! CY搬出時間 -->
<% if F_CYDelTime(1) = "1" or Request.Form("F_CYDelTime2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="CYDelTime2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="CYDelTime2"></TD>
<% end if %>

<! ディテンションフリータイム -->
<% if F_DetentionFreeTime(1) = "1" or Request.Form("F_DetentionFreeTime2") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DetentionFreeTime2" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DetentionFreeTime2"></TD>
<% end if %>

<! 空コン返却状況 -->
<% if F_ReturnTime(1) = "1" or Request.Form("F_ReturnTime2") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ReturnTime2" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ReturnTime2"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-E -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
<% if Request.Form("Email3") <> "" then %>
		<TD width="95%" colspan="2">3　<input type="text" name="Email3" value="<%=Request.Form("Email3")%>" size="70" maxlength="100"></TD>
<% else %>
		<TD width="95%" colspan="2">3　<input type="text" name="Email3" value="<%=Email3%>" size="70" maxlength="100"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-S -->
<! 入港時間 -->
<% if F_ArrivalTime(2) = "1" or Request.Form("F_ArrivalTime3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ArrivalTime3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ArrivalTime3"></TD>
<% end if %>

<! CY搬入時間 -->
<% if F_InTime(2) = "1" or Request.Form("F_InTime3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="InTime3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="InTime3"></TD>
<% end if %>

<! 通関許可状況 -->
<% if F_List(2) = "1" or Request.Form("F_List3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="List3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="List3"></TD>
<% end if %>

<! ＤＯクリア状況 -->
<% if F_DOStatus(2) = "1" or Request.Form("F_DOStatus3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DOStatus3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DOStatus3"></TD>
<% end if %>

<! 搬出可否 -->
<% if F_DelPermit(2) = "1" or Request.Form("F_DelPermit3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DelPermit3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DelPermit3"></TD>
<% end if %>

<! デマレージフリータイム -->
<% if F_DemurrageFreeTime(2) = "1" or Request.Form("F_DemurrageFreeTime3") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DemurrageFreeTime3" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DemurrageFreeTime3"></TD>
<% end if %>

<! CY搬出時間 -->
<% if F_CYDelTime(2) = "1" or Request.Form("F_CYDelTime3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="CYDelTime3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="CYDelTime3"></TD>
<% end if %>

<! ディテンションフリータイム -->
<% if F_DetentionFreeTime(2) = "1" or Request.Form("F_DetentionFreeTime3") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DetentionFreeTime3" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DetentionFreeTime3"></TD>
<% end if %>

<! 空コン返却状況 -->
<% if F_ReturnTime(2) = "1" or Request.Form("F_ReturnTime3") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ReturnTime3" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ReturnTime3"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-E -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
<% if Request.Form("Email4") <> "" then %>
		<TD width="95%" colspan="2">4　<input type="text" name="Email4" value="<%=Request.Form("Email4")%>" size="70" maxlength="100"></TD>
<% else %>
		<TD width="95%" colspan="2">4　<input type="text" name="Email4" value="<%=Email4%>" size="70" maxlength="100"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-S -->
<! 入港時間 -->
<% if F_ArrivalTime(3) = "1" or Request.Form("F_ArrivalTime4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ArrivalTime4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ArrivalTime4"></TD>
<% end if %>

<! CY搬入時間 -->
<% if F_InTime(3) = "1" or Request.Form("F_InTime4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="InTime4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="InTime4"></TD>
<% end if %>

<! 通関許可状況 -->
<% if F_List(3) = "1" or Request.Form("F_List4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="List4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="List4"></TD>
<% end if %>

<! ＤＯクリア状況 -->
<% if F_DOStatus(3) = "1" or Request.Form("F_DOStatus4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DOStatus4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DOStatus4"></TD>
<% end if %>

<! 搬出可否 -->
<% if F_DelPermit(3) = "1" or Request.Form("F_DelPermit4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DelPermit4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DelPermit4"></TD>
<% end if %>

<! デマレージフリータイム -->
<% if F_DemurrageFreeTime(3) = "1" or Request.Form("F_DemurrageFreeTime4") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DemurrageFreeTime4" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DemurrageFreeTime4"></TD>
<% end if %>

<! CY搬出時間 -->
<% if F_CYDelTime(3) = "1" or Request.Form("F_CYDelTime4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="CYDelTime4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="CYDelTime4"></TD>
<% end if %>

<! ディテンションフリータイム -->
<% if F_DetentionFreeTime(3) = "1" or Request.Form("F_DetentionFreeTime4") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DetentionFreeTime4" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DetentionFreeTime4"></TD>
<% end if %>

<! 空コン返却状況 -->
<% if F_ReturnTime(3) = "1" or Request.Form("F_ReturnTime4") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ReturnTime4" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ReturnTime4"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-E -->
	</TR>
	<TR>
		<TD width="5%">　</TD>
<% if Request.Form("Email5") <> "" then %>
		<TD width="95%" colspan="2">5　<input type="text" name="Email5" value="<%=Request.Form("Email5")%>" size="70" maxlength="100"></TD>
<% else %>
		<TD width="95%" colspan="2">5　<input type="text" name="Email5" value="<%=Email5%>" size="70" maxlength="100"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-S -->
<! 入港時間 -->
<% if F_ArrivalTime(4) = "1" or Request.Form("F_ArrivalTime5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ArrivalTime5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ArrivalTime5"></TD>
<% end if %>

<! CY搬入時間 -->
<% if F_InTime(4) = "1" or Request.Form("F_InTime5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="InTime5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="InTime5"></TD>
<% end if %>

<! 通関許可状況 -->
<% if F_List(4) = "1" or Request.Form("F_List5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="List5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="List5"></TD>
<% end if %>

<! ＤＯクリア状況 -->
<% if F_DOStatus(4) = "1" or Request.Form("F_DOStatus5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DOStatus5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DOStatus5"></TD>
<% end if %>

<! 搬出可否 -->
<% if F_DelPermit(4) = "1" or Request.Form("F_DelPermit5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="DelPermit5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="DelPermit5"></TD>
<% end if %>

<! デマレージフリータイム -->
<% if F_DemurrageFreeTime(4) = "1" or Request.Form("F_DemurrageFreeTime5") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DemurrageFreeTime5" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DemurrageFreeTime5"></TD>
<% end if %>

<! CY搬出時間 -->
<% if F_CYDelTime(4) = "1" or Request.Form("F_CYDelTime5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="CYDelTime5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="CYDelTime5"></TD>
<% end if %>

<! ディテンションフリータイム -->
<% if F_DetentionFreeTime(4) = "1" or Request.Form("F_DetentionFreeTime5") = "1" then %>
		<TD width="5%"><input type="checkbox" value="1" name="DetentionFreeTime5" checked></TD>
<% else %>
		<TD width="5%"><input type="checkbox" value="0" name="DetentionFreeTime5"></TD>
<% end if %>

<! //空コン返却状況 -->
<% if F_ReturnTime(4) = "1" or Request.Form("F_ReturnTime5") = "1" then %>
		<TD width="5%" colspan="1"><input type="checkbox" value="1" name="ReturnTime5" checked></TD>
<% else %>
		<TD width="5%" colspan="1"><input type="checkbox" value="0" name="ReturnTime5"></TD>
<% end if %>
<!--	2009/03/10 R.Shibuta Add-E -->
	</TR>
	<TR>
		<TD colspan="3" align="center">
<!--	2009/03/10 R.Shibuta Upd-S -->
<!--		<INPUT type="hidden" name="F_ArrivalTime" value=""> -->
<!--		<INPUT type="hidden" name="F_InTime" value=""> -->
<!--		<INPUT type="hidden" name="F_List" value=""> -->
<!--		<INPUT type="hidden" name="F_DOStatus" value=""> -->
<!--		<INPUT type="hidden" name="F_DelPermit" value=""> -->
<!--		<INPUT type="hidden" name="F_DemurrageFreeTime" value=""> -->
<!--		<INPUT type="hidden" name="F_CYDelTime" value=""> -->
<!--		<INPUT type="hidden" name="F_DetentionFreeTime" value=""> -->
<!--		<INPUT type="hidden" name="F_ReturnTime" value=""> -->
<!--		<INPUT type="submit" value="登録" onClick="return GoEntry()"> -->
<!--		<INPUT type="submit" value="中止" onClick="GoStop()"> -->
			<INPUT type="hidden" name="F_ArrivalTime1" value="">
			<INPUT type="hidden" name="F_ArrivalTime2" value="">
			<INPUT type="hidden" name="F_ArrivalTime3" value="">
			<INPUT type="hidden" name="F_ArrivalTime4" value="">
			<INPUT type="hidden" name="F_ArrivalTime5" value="">
			<INPUT type="hidden" name="F_InTime1" value="">
			<INPUT type="hidden" name="F_InTime2" value="">
			<INPUT type="hidden" name="F_InTime3" value="">
			<INPUT type="hidden" name="F_InTime4" value="">
			<INPUT type="hidden" name="F_InTime5" value="">
			<INPUT type="hidden" name="F_List1" value="">
			<INPUT type="hidden" name="F_List2" value="">
			<INPUT type="hidden" name="F_List3" value="">
			<INPUT type="hidden" name="F_List4" value="">
			<INPUT type="hidden" name="F_List5" value="">
			<INPUT type="hidden" name="F_DOStatus1" value="">
			<INPUT type="hidden" name="F_DOStatus2" value="">
			<INPUT type="hidden" name="F_DOStatus3" value="">
			<INPUT type="hidden" name="F_DOStatus4" value="">
			<INPUT type="hidden" name="F_DOStatus5" value="">
			<INPUT type="hidden" name="F_DelPermit1" value="">
			<INPUT type="hidden" name="F_DelPermit2" value="">
			<INPUT type="hidden" name="F_DelPermit3" value="">
			<INPUT type="hidden" name="F_DelPermit4" value="">
			<INPUT type="hidden" name="F_DelPermit5" value="">
			<INPUT type="hidden" name="F_DemurrageFreeTime1" value="">
			<INPUT type="hidden" name="F_DemurrageFreeTime2" value="">
			<INPUT type="hidden" name="F_DemurrageFreeTime3" value="">
			<INPUT type="hidden" name="F_DemurrageFreeTime4" value="">
			<INPUT type="hidden" name="F_DemurrageFreeTime5" value="">
			<INPUT type="hidden" name="F_CYDelTime1" value="">
			<INPUT type="hidden" name="F_CYDelTime2" value="">
			<INPUT type="hidden" name="F_CYDelTime3" value="">
			<INPUT type="hidden" name="F_CYDelTime4" value="">
			<INPUT type="hidden" name="F_CYDelTime5" value="">
			<INPUT type="hidden" name="F_DetentionFreeTime1" value="">
			<INPUT type="hidden" name="F_DetentionFreeTime2" value="">
			<INPUT type="hidden" name="F_DetentionFreeTime3" value="">
			<INPUT type="hidden" name="F_DetentionFreeTime4" value="">
			<INPUT type="hidden" name="F_DetentionFreeTime5" value="">
			<INPUT type="hidden" name="F_ReturnTime1" value="">
			<INPUT type="hidden" name="F_ReturnTime2" value="">
			<INPUT type="hidden" name="F_ReturnTime3" value="">
			<INPUT type="hidden" name="F_ReturnTime4" value="">
			<INPUT type="hidden" name="F_ReturnTime5" value="">
			<INPUT type="submit" value="REGISTER" onClick="return GoEntry()">
			<INPUT type="submit" value="CANCEL" onClick="GoStop()">
<!--	2009/03/10 R.Shibuta Upd-E -->
		</TD>
	</TR>
</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
