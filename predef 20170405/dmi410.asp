<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits								    	   _/
'_/	FileName	:dmi410.asp									   _/
'_/	Function	:作業発生mail対象項目設定					   _/
'_/	Date			:2009/03/10								   _/
'_/	Code By		:Shbuta    									   _/
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
	Dim F_DelResults(4), F_RecEmp(4), F_RecResults(4), F_DelEmp(4)
	Dim Email1, Email2, Email3, Email4, Email5
	Dim iCnt
	Dim USER

 	'入力内容の確認画面からの戻りでない場合、
	'すなわち「設定」メニューから来た場合にのみＤＢから値を取得する。
	if Session.Contents("dmi411") <> "true" then

		USER = Session.Contents("userid")
	  
	 	'DB接続
	 	Dim ObjConn, ObjRS, StrSQL
	 	ConnDBH ObjConn, ObjRS
	 
	 	StrSQL = "SELECT * from TargetOperation where UserCode='"& USER &"'"
	 	ObjRS.Open StrSQL, ObjConn
	 
	 	if err <> 0 then
	 		DisConnDBH ObjConn, ObjRS	'DB切断
	 		jumpErrorP "1","c103","01","作業発生mail設定","101","SQL:<BR>"&strSQL
	 	end if
	 	
	 	if ObjRS.eof then
	 		for iCnt = 0 To 4
	 			F_DelResults(iCnt) = ""
	 			F_RecEmp(iCnt) = ""
	 			F_RecResults(iCnt) = ""
	 			F_DelEmp(iCnt) = ""
	 		next
	 		Email1 = ""
	 		Email2 = ""
	 		Email3 = ""
	 		Email4 = ""
	 		Email5 = ""
	 	else
	 		F_DelResults(0) = ObjRS("FlagDelResults1")
	 		F_DelResults(1) = ObjRS("FlagDelResults2")
	 		F_DelResults(2) = ObjRS("FlagDelResults3")
	 		F_DelResults(3) = ObjRS("FlagDelResults4")
	 		F_DelResults(4) = ObjRS("FlagDelResults5")
	 		
	 		F_RecEmp(0) = ObjRS("FlagRecEmp1")
	 		F_RecEmp(1) = ObjRS("FlagRecEmp2")
	 		F_RecEmp(2) = ObjRS("FlagRecEmp3")
	 		F_RecEmp(3) = ObjRS("FlagRecEmp4")
	 		F_RecEmp(4) = ObjRS("FlagRecEmp5")
	 		
	 		F_RecResults(0) = ObjRS("FlagRecResults1")
	 		F_RecResults(1) = ObjRS("FlagRecResults2")
	 		F_RecResults(2) = ObjRS("FlagRecResults3")
	 		F_RecResults(3) = ObjRS("FlagRecResults4")
	 		F_RecResults(4) = ObjRS("FlagRecResults5")
	 		
	 		F_DelEmp(0) = ObjRS("FlagDelEmp1")
	 		F_DelEmp(1) = ObjRS("FlagDelEmp2")
	 		F_DelEmp(2) = ObjRS("FlagDelEmp3")
	 		F_DelEmp(3) = ObjRS("FlagDelEmp4")
	 		F_DelEmp(4) = ObjRS("FlagDelEmp5")
	 		
	 		Email1 = Trim(ObjRS("Email1"))
	 		Email2 = Trim(ObjRS("Email2"))
	 		Email3 = Trim(ObjRS("Email3"))
	 		Email4 = Trim(ObjRS("Email4"))
	 		Email5 = Trim(ObjRS("Email5"))
		end if
	 	
	 	ObjRS.close
	 	
	 	'DB接続解除
	 		DisConnDBH ObjConn, ObjRS
	 	'エラートラップ解除
	 		on error goto 0
	 	'ログ出力
	 	 WriteLogH "b402", "作業発生mail設定","00",""
	 end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>作業発生mail対象項目設定</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
function GoEntry(){

	f=document.dmi410;
	//メールアドレスの内容チェック
	if(document.getElementById("Email1").value!=""){
		if(gfisMailAddr(document.getElementById("Email1").value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			document.getElementById("Email1").focus();
			return false;
		}
		if(document.getElementById("Email1").value==document.getElementById("Email2").value || document.getElementById("Email1").value==document.getElementById("Email3").value ||
			 document.getElementById("Email1").value==document.getElementById("Email4").value || document.getElementById("Email1").value==document.getElementById("Email5").value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				document.getElementById("Email1").focus();
				return false;
			}
		}
	}
	if(document.getElementById("Email2").value!=""){
		if(gfisMailAddr(document.getElementById("Email2").value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			document.getElementById("Email2").focus();
			return false;
		}
		if(document.getElementById("Email2").value==document.getElementById("Email3").value || document.getElementById("Email2").value==document.getElementById("Email4").value ||
			 document.getElementById("Email2").value==document.getElementById("Email5").value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				document.getElementById("Email2").focus();
			return false;
			}
		}
	}
	if(document.getElementById("Email3").value!=""){
		if(gfisMailAddr(document.getElementById("Email3").value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			document.getElementById("Email3").focus();
			return false;
		}
		if(document.getElementById("Email3").value==document.getElementById("Email4").value || document.getElementById("Email3").value==document.getElementById("Email5").value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				document.getElementById("Email3").focus();
				return false;
			}
		}
	}
	if(document.getElementById("Email4").value!=""){
		if(gfisMailAddr(document.getElementById("Email4").value)==false){
			alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
			document.getElementById("Email4").focus();
			return false;
		}
		if(document.getElementById("Email4").value==document.getElementById("Email5").value){
			if(!confirm("同じメールアドレスが指定されています。\nこのまま登録してよろしいですか？")){
				document.getElementById("Email4").focus();
				return false;
			}
		}
	}
	if((document.getElementById("Email5").value!="") && (gfisMailAddr(document.getElementById("Email5").value)==false)){
		alert("メールアドレスが不正です。\nメールアドレスを確認してください。");
		document.getElementById("Email5").focus();
		return false;
	}
		
	if(document.getElementById("DelResults1").checked==true){
		document.getElementById("F_DelResults1").value="1"
	}else{
		document.getElementById("F_DelResults1").value="0"
	}
	
	if(document.getElementById("DelResults2").checked==true){
		document.getElementById("F_DelResults2").value="1"
	}else{
		document.getElementById("F_DelResults2").value="0"
	}
	if(document.getElementById("DelResults3").checked==true){
		document.getElementById("F_DelResults3").value="1"
	}else{
		document.getElementById("F_DelResults3").value="0"
	}
	if(document.getElementById("DelResults4").checked==true){
		document.getElementById("F_DelResults4").value="1"
	}else{
		document.getElementById("F_DelResults4").value="0"
	}
	if(document.getElementById("DelResults5").checked==true){
		document.getElementById("F_DelResults5").value="1"
	}else{
		document.getElementById("F_DelResults5").value="0"
	}
		
	if(document.getElementById("RecEmp1").checked==true){
		document.getElementById("F_RecEmp1").value="1"
	}else{
		document.getElementById("F_RecEmp1").value="0"
	}
	if(document.getElementById("RecEmp2").checked==true){
		document.getElementById("F_RecEmp2").value="1"
	}else{
		document.getElementById("F_RecEmp2").value="0"
	}
	if(document.getElementById("RecEmp3").checked==true){
		document.getElementById("F_RecEmp3").value="1"
	}else{
		document.getElementById("F_RecEmp3").value="0"
	}
	if(document.getElementById("RecEmp4").checked==true){
		document.getElementById("F_RecEmp4").value="1"
	}else{
		document.getElementById("F_RecEmp4").value="0"
	}
	if(document.getElementById("RecEmp5").checked==true){
		document.getElementById("F_RecEmp5").value="1"
	}else{
		document.getElementById("F_RecEmp5").value="0"
	}

	if(document.getElementById("RecResults1").checked==true){
		document.getElementById("F_RecResults1").value="1"
	}else{
		document.getElementById("F_RecResults1").value="0"
	}
	if(document.getElementById("RecResults2").checked==true){
		document.getElementById("F_RecResults2").value="1"
	}else{
		document.getElementById("F_RecResults2").value="0"
	}
	if(document.getElementById("RecResults3").checked==true){
		document.getElementById("F_RecResults3").value="1"
	}else{
		document.getElementById("F_RecResults3").value="0"
	}
	if(document.getElementById("RecResults4").checked==true){
		document.getElementById("F_RecResults4").value="1"
	}else{
		document.getElementById("F_RecResults4").value="0"
	}
	if(document.getElementById("RecResults5").checked==true){
		document.getElementById("F_RecResults5").value="1"
	}else{
		document.getElementById("F_RecResults5").value="0"
	}
	
	if(document.getElementById("DelEmp1").checked==true){
		document.getElementById("F_DelEmp1").value="1"
	}else{
		document.getElementById("F_DelEmp1").value="0"
	}
	if(document.getElementById("DelEmp2").checked==true){
		document.getElementById("F_DelEmp2").value="1"
	}else{
		document.getElementById("F_DelEmp2").value="0"
	}
	if(document.getElementById("DelEmp3").checked==true){
		document.getElementById("F_DelEmp3").value="1"
	}else{
		document.getElementById("F_DelEmp3").value="0"
	}
	if(document.getElementById("DelEmp4").checked==true){
		document.getElementById("F_DelEmp4").value="1"
	}else{
		document.getElementById("F_DelEmp4").value="0"
	}
	if(document.getElementById("DelEmp5").checked==true){
		document.getElementById("F_DelEmp5").value="1"
	}else{
		document.getElementById("F_DelEmp5").value="0"
	}
	
	f.action="dmi411.asp";
	return true;
}

function GoStop(){
<%	Session.Contents("dmi411") = "false" %>
	window.close();
}

function GoClear(){

document.getElementById("DelResults1").checked = false;
document.getElementById("DelResults2").checked = false;
document.getElementById("DelResults3").checked = false;
document.getElementById("DelResults4").checked = false;
document.getElementById("DelResults5").checked = false;

document.getElementById("RecEmp1").checked = false;
document.getElementById("RecEmp2").checked = false;
document.getElementById("RecEmp3").checked = false;
document.getElementById("RecEmp4").checked = false;
document.getElementById("RecEmp5").checked = false;

document.getElementById("RecResults1").checked = false;
document.getElementById("RecResults2").checked = false;
document.getElementById("RecResults3").checked = false;
document.getElementById("RecResults4").checked = false;
document.getElementById("RecResults5").checked = false;

document.getElementById("DelEmp1").checked = false;
document.getElementById("DelEmp2").checked = false;
document.getElementById("DelEmp3").checked = false;
document.getElementById("DelEmp4").checked = false;
document.getElementById("DelEmp5").checked = false;

document.getElementById("Email1").value ='';
document.getElementById("Email2").value ='';
document.getElementById("Email3").value ='';
document.getElementById("Email4").value ='';
document.getElementById("Email5").value ='';

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

</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------作業発生mail項目設定画面--------------------------->
<%'データ登録／更新しました画面にて「最新の情報に更新」でSubmitされた場合の対策 %>
<% Session.Contents("ItemsSubmitted")="False"  %>
<FORM name="dmi410" method="POST">
<TABLE border="0" cellPadding="5" cellSpacing="0" width="100%">

	<TR><TD>　</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">　●自分宛の作業依頼が発生した場合にmailで連絡します。<BR>　　下でmail送信先とともに希望する作業の番号にチェックしてください。</TD>
	</TR>
	
	<TR>
		<TD width="5%">　</TD>
		<TD width="40%">　　　（１）実搬出作業<TD>
		<TD width="55%">（２）空搬入作業<TD>
	</TR>
	
	<TR>
		<TD width="5%">　</TD>
		<TD width="40%">　　　（３）実搬入作業<TD>
		<TD width="55%">（４）空搬出作業<TD>
	</TR>
	
	<TR><TD>　</TD></TR>

	<TR>
		<TD width="5%" colspan="4" >　●mailの送信先を設定してください。</TD>
		<TD width="1%" align="center">(1)</TD>
		<TD width="1%" align="center">(2)</TD>
		<TD width="1%" align="center">(3)</TD>
		<TD width="1%" align="center">(4)</TD>
	</TR>
	
	<TR>
		<TD width="5%">　</TD>
		<% if Request.Form("Email1") <> "" then %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email1" value="<%=Request.Form("Email1")%>" size="70" maxlength="100"></TD>
		<% else %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email1" value="<%=Email1%>" size="70" maxlength="100"></TD>
		<% end if %>
		
		<% if F_DelResults(0) = "1" or Request.Form("F_DelResults1") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelResults1" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelResults1"></TD>
		<% end if %>
		
		<% if F_RecEmp(0) = "1" or Request.Form("F_RecEmp1") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecEmp1" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecEmp1"></TD>
		<% end if %>

		<% if F_RecResults(0) = "1" or Request.Form("F_RecResults1") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecResults1" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecResults1"></TD>
		<% end if %>

		<% if F_DelEmp(0) = "1" or Request.Form("F_DelEmp1") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelEmp1" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelEmp1"></TD>
		<% end if %>

	</TR>

	<TR>
		<TD width="5%">　</TD>
		<% if Request.Form("Email2") <> "" then %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email2" value="<%=Request.Form("Email2")%>" size="70" maxlength="100"></TD>
		<% else %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email2" value="<%=Email2%>" size="70" maxlength="100"></TD>
		<% end if %>
		
		<% if F_DelResults(1) = "1" or Request.Form("F_DelResults2") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelResults2" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelResults2"></TD>
		<% end if %>
		
		<% if F_RecEmp(1) = "1" or Request.Form("F_RecEmp2") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecEmp2" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecEmp2"></TD>
		<% end if %>

		<% if F_RecResults(1) = "1" or Request.Form("F_RecResults2") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecResults2" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecResults2"></TD>
		<% end if %>

		<% if F_DelEmp(1) = "1" or Request.Form("F_DelEmp2") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelEmp2" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelEmp2"></TD>
		<% end if %>
		
	</TR>

	<TR>
		<TD width="5%">　</TD>
		<% if Request.Form("Email3") <> "" then %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email3" value="<%=Request.Form("Email3")%>" size="70" maxlength="100"></TD>
		<% else %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email3" value="<%=Email3%>" size="70" maxlength="100"></TD>
		<% end if %>
		
		<% if F_DelResults(2) = "1" or Request.Form("F_DelResults3") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelResults3" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelResults3"></TD>
		<% end if %>
		
		<% if F_RecEmp(2) = "1" or Request.Form("F_RecEmp3") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecEmp3" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecEmp3"></TD>
		<% end if %>

		<% if F_RecResults(2) = "1" or Request.Form("F_RecResults3") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecResults3" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecResults3"></TD>
		<% end if %>

		<% if F_DelEmp(2) = "1" or Request.Form("F_DelEmp3") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelEmp3" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelEmp3"></TD>
		<% end if %>
		
	</TR>

	<TR>
		<TD width="5%">　</TD>
		<% if Request.Form("Email4") <> "" then %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email4" value="<%=Request.Form("Email4")%>" size="70" maxlength="100"></TD>
		<% else %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email4" value="<%=Email4%>" size="70" maxlength="100"></TD>
		<% end if %>
		
		<% if F_DelResults(3) = "1" or Request.Form("F_DelResults4") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelResults4" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelResults4"></TD>
		<% end if %>
		
		<% if F_RecEmp(3) = "1" or Request.Form("F_RecEmp4") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecEmp4" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecEmp4"></TD>
		<% end if %>

		<% if F_RecResults(3) = "1" or Request.Form("F_RecResults4") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecResults4" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecResults4"></TD>
		<% end if %>

		<% if F_DelEmp(3) = "1" or Request.Form("F_DelEmp4") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelEmp4" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelEmp4"></TD>
		<% end if %>
		
	</TR>

	<TR>
		<TD width="5%">　</TD>
		<% if Request.Form("Email5") <> "" then %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email5" value="<%=Request.Form("Email5")%>" size="70" maxlength="100"></TD>
		<% else %>
				<TD width="5%" colspan="3">　　<input type="text" name="Email5" value="<%=Email5%>" size="70" maxlength="100"></TD>
		<% end if %>
		
		<% if F_DelResults(4) = "1" or Request.Form("F_DelResults5") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelResults5" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelResults5"></TD>
		<% end if %>
		
		<% if F_RecEmp(4) = "1" or Request.Form("F_RecEmp5") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecEmp5" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecEmp5"></TD>
		<% end if %>

		<% if F_RecResults(4) = "1" or Request.Form("F_RecResults5") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="RecResults5" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="RecResults5"></TD>
		<% end if %>

		<% if F_DelEmp(4) = "1" or Request.Form("F_DelEmp5") = "1" then %>
				<TD width="1%" colspan="1"><input type="checkbox" value="1" name="DelEmp5" checked></TD>
		<% else %>
				<TD width="1%" colspan="1"><input type="checkbox" value="0" name="DelEmp5"></TD>
		<% end if %>
		
	</TR>

	<TR>
		<TD colspan="5" align="center">
			<INPUT type="hidden" name="F_DelResults1" value="">
			<INPUT type="hidden" name="F_DelResults2" value="">
			<INPUT type="hidden" name="F_DelResults3" value="">
			<INPUT type="hidden" name="F_DelResults4" value="">
			<INPUT type="hidden" name="F_DelResults5" value="">
			<INPUT type="hidden" name="F_RecEmp1" value="">
			<INPUT type="hidden" name="F_RecEmp2" value="">
			<INPUT type="hidden" name="F_RecEmp3" value="">
			<INPUT type="hidden" name="F_RecEmp4" value="">
			<INPUT type="hidden" name="F_RecEmp5" value="">
			<INPUT type="hidden" name="F_RecResults1" value="">
			<INPUT type="hidden" name="F_RecResults2" value="">
			<INPUT type="hidden" name="F_RecResults3" value="">
			<INPUT type="hidden" name="F_RecResults4" value="">
			<INPUT type="hidden" name="F_RecResults5" value="">
			<INPUT type="hidden" name="F_DelEmp1" value="">
			<INPUT type="hidden" name="F_DelEmp2" value="">
			<INPUT type="hidden" name="F_DelEmp3" value="">
			<INPUT type="hidden" name="F_DelEmp4" value="">
			<INPUT type="hidden" name="F_DelEmp5" value="">
			<INPUT type="submit" value="登録" onClick="return GoEntry()">
			<INPUT type="submit" value="中止" onClick="GoStop()">
		    <INPUT type="button" value="クリア" onClick="GoClear()">
		</TD>
	</TR>
</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
