<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										    _/
'_/	FileName	:dmi411.asp									    _/
'_/	Function	:作業発生mail対象項目設定入力確認			    _/
'_/	Date		:2009/03/10									    _/
'_/	Code By		:Shibuta									    _/
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
	Dim NoEntered, ItemsToSend
	DIm strWork
	
	F_DelResults(0) = Request.Form("F_DelResults1")
	F_DelResults(1) = Request.Form("F_DelResults2")
	F_DelResults(2) = Request.Form("F_DelResults3")
	F_DelResults(3) = Request.Form("F_DelResults4")
	F_DelResults(4) = Request.Form("F_DelResults5")
	
	F_RecEmp(0) = Request.Form("F_RecEmp1")
	F_RecEmp(1) = Request.Form("F_RecEmp2")
	F_RecEmp(2) = Request.Form("F_RecEmp3")
	F_RecEmp(3) = Request.Form("F_RecEmp4")
	F_RecEmp(4) = Request.Form("F_RecEmp5")
	
	F_RecResults(0) = Request.Form("F_RecResults1")
	F_RecResults(1) = Request.Form("F_RecResults2")
	F_RecResults(2) = Request.Form("F_RecResults3")
	F_RecResults(3) = Request.Form("F_RecResults4")
	F_RecResults(4) = Request.Form("F_RecResults5")
	
	F_DelEmp(0) = Request.Form("F_DelEmp1")
	F_DelEmp(1) = Request.Form("F_DelEmp2")
	F_DelEmp(2) = Request.Form("F_DelEmp3")
	F_DelEmp(3) = Request.Form("F_DelEmp4")
	F_DelEmp(4) = Request.Form("F_DelEmp5")
	
	Email1 = Request.Form("Email1")
	Email2 = Request.Form("Email2")
	Email3 = Request.Form("Email3")
	Email4 = Request.Form("Email4")
	Email5 = Request.Form("Email5")
 	
	Session.Contents("dmi411") = "true"
	
 	'何も入力されていない場合
	For iCnt = 0 To 4
		if F_DelResults(iCnt) = "0" and F_RecEmp(iCnt) = "0" and F_RecResults(iCnt) = "0" and F_DelEmp(iCnt) = "0" _
			and Email1 = "" and Email2 = "" and Email3 = "" and Email4 = "" and Email5 = "" then
			NoEntered = "1"
		else
			NoEntered = "0"
		end if
		if NoEntered = "0" then
			Exit For
		end if
	Next
 	
 	'メール送信対象項目数
	ItemsToSend = 0
 	
 	'ログ出力
 	WriteLogH "c402", "作業発生mail設定","02",""


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>作業発生mail対象項目設定入力確認</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>

//登録
function GoEntry(){
	f=document.dmi411;
	f.action="dmi412.asp";
	return true;
}

//戻る
function GoBack(){
	f=document.dmi411;
	f.action="dmi410.asp";
	return true;
}

</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------ステータス配信対象項目設定画面--------------------------->
<%'データ登録／更新しました画面にて「最新の情報に更新」でSubmitされた場合の対策 %>
<% if NoEntered = "0" then %>
<% Session.Contents("ItemsSubmitted")="False"  %>
<FORM name="dmi411" method="POST">
<TABLE border="0" cellPadding="5" cellSpacing="0" width="100%">

	<TR>
		<TD width="5%" colspan="20">　<B>作業発生mail（設定）項目確認</B></TD>
	</TR>
	
	<TR><TD>　</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">　以下の作業依頼が発生した場合にmailで連絡します。</TD>
	</TR>

	<% For iCnt = 0 To 4 %>
		<% if F_DelResults(iCnt) = "1" then %>
			<TR>
				<TD width="40%">　　（１）実搬出作業<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>  
		
	<% For iCnt = 0 To 4 %>
		<% if F_RecEmp(iCnt) = "1" then %>
			<TR>
				<TD width="40%">　　（２）空搬入作業<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>
	
	<% For iCnt = 0 To 4 %>
		<% if F_RecResults(iCnt) = "1" then %>
			<TR>
				<TD width="40%">　　（３）実搬入作業<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>
	
	<% For iCnt = 0 To 4 %>
		<% if F_DelEmp(iCnt) = "1" then %>
			<TR>
				<TD width="40%">　　（４）空搬出作業<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>
	
	<TR><TD>　</TD></TR>

	<TR>
		<TD width="20%">　●送信先</TD>
	</TR>
	
<% if Email1 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">　　<%=Email1%></TD>
		
		<% if F_DelResults(0) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(0) = "1" then %>
			<% if F_DelResults(0) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(0) = "1" then %>
			<% if F_DelResults(0) = "1" Or F_RecEmp(0) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(0) = "1" then %>
			<% if F_DelResults(0) = "1" Or F_RecEmp(0) = "1" Or F_RecResults(0) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email2 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">　　<%=Email2%></TD>
		
		<% if F_DelResults(1) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(1) = "1" then %>
			<% if F_DelResults(1) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(1) = "1" then %>
			<% if F_DelResults(1) = "1" Or F_RecEmp(1) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(1) = "1" then %>
			<% if F_DelResults(1) = "1" Or F_RecEmp(1) = "1" Or F_RecResults(1) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email3 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">　　<%=Email3%></TD>
		
		<% if F_DelResults(2) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(2) = "1" then %>
			<% if F_DelResults(2) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(2) = "1" then %>
			<% if F_DelResults(2) = "1" Or F_RecEmp(2) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(2) = "1" then %>
			<% if F_DelResults(2) = "1" Or F_RecEmp(2) = "1" Or F_RecResults(2) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email4 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">　　<%=Email4%></TD>
		
		<% if F_DelResults(3) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(3) = "1" then %>
			<% if F_DelResults(3) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(3) = "1" then %>
			<% if F_DelResults(3) = "1" Or F_RecEmp(3) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(3) = "1" then %>
			<% if F_DelResults(3) = "1" Or F_RecEmp(3) = "1" Or F_RecResults(3) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email5 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">　　<%=Email5%></TD>
		
		<% if F_DelResults(4) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(4) = "1" then %>
			<% if F_DelResults(4) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(4) = "1" then %>
			<% if F_DelResults(4) = "1" Or F_RecEmp(4) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(4) = "1" then %>
			<% if F_DelResults(4) = "1" Or F_RecEmp(4) = "1" Or F_RecResults(4) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

	<TR>
		<TD colspan="5" align="center">
			<INPUT type="hidden" name="F_DelResults1" value="<%=F_DelResults(0)%>">
			<INPUT type="hidden" name="F_DelResults2" value="<%=F_DelResults(1)%>">
			<INPUT type="hidden" name="F_DelResults3" value="<%=F_DelResults(2)%>">
			<INPUT type="hidden" name="F_DelResults4" value="<%=F_DelResults(3)%>">
			<INPUT type="hidden" name="F_DelResults5" value="<%=F_DelResults(4)%>">
			
			<INPUT type="hidden" name="F_RecEmp1" value="<%=F_RecEmp(0)%>">
			<INPUT type="hidden" name="F_RecEmp2" value="<%=F_RecEmp(1)%>">
			<INPUT type="hidden" name="F_RecEmp3" value="<%=F_RecEmp(2)%>">
			<INPUT type="hidden" name="F_RecEmp4" value="<%=F_RecEmp(3)%>">
			<INPUT type="hidden" name="F_RecEmp5" value="<%=F_RecEmp(4)%>">
			
			<INPUT type="hidden" name="F_RecResults1" value="<%=F_RecResults(0)%>">
			<INPUT type="hidden" name="F_RecResults2" value="<%=F_RecResults(1)%>">
			<INPUT type="hidden" name="F_RecResults3" value="<%=F_RecResults(2)%>">
			<INPUT type="hidden" name="F_RecResults4" value="<%=F_RecResults(3)%>">
			<INPUT type="hidden" name="F_RecResults5" value="<%=F_RecResults(4)%>">
			
			<INPUT type="hidden" name="F_DelEmp1" value="<%=F_DelEmp(0)%>">
			<INPUT type="hidden" name="F_DelEmp2" value="<%=F_DelEmp(1)%>">
			<INPUT type="hidden" name="F_DelEmp3" value="<%=F_DelEmp(2)%>">
			<INPUT type="hidden" name="F_DelEmp4" value="<%=F_DelEmp(3)%>">
			<INPUT type="hidden" name="F_DelEmp5" value="<%=F_DelEmp(4)%>">
			
			<INPUT type="hidden" name="Email1" value="<%=Email1%>">
			<INPUT type="hidden" name="Email2" value="<%=Email2%>">
			<INPUT type="hidden" name="Email3" value="<%=Email3%>">
			<INPUT type="hidden" name="Email4" value="<%=Email4%>">
			<INPUT type="hidden" name="Email5" value="<%=Email5%>">
			
			<INPUT type="submit" value="ＯＫ" onClick="return GoEntry()">
			<INPUT type="submit" value="戻る" onClick="GoBack()">
		</TD>
	</TR>  
</TABLE>
</FORM>
<% else %>
<FORM name="dmi411" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
	<TR><TD>　</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">　<B>作業発生mail（設定）項目確認</B></TD>
	</TR>
	
	<TR><TD>　</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">　何も指定されていません。よろしければ「ＯＫ」ボタンをクリックしてください。</TD>
	</TR>

	<TR><TD>　</TD></TR>
	<TR>
		<TD colspan="5" align="center">
			<INPUT type="hidden" name="F_DelResults1" value="<%=F_DelResults(0)%>">
			<INPUT type="hidden" name="F_DelResults2" value="<%=F_DelResults(1)%>">
			<INPUT type="hidden" name="F_DelResults3" value="<%=F_DelResults(2)%>">
			<INPUT type="hidden" name="F_DelResults4" value="<%=F_DelResults(3)%>">
			<INPUT type="hidden" name="F_DelResults5" value="<%=F_DelResults(4)%>">

			<INPUT type="hidden" name="F_RecEmp1" value="<%=F_RecEmp(0)%>">
			<INPUT type="hidden" name="F_RecEmp2" value="<%=F_RecEmp(1)%>">
			<INPUT type="hidden" name="F_RecEmp3" value="<%=F_RecEmp(2)%>">
			<INPUT type="hidden" name="F_RecEmp4" value="<%=F_RecEmp(3)%>">
			<INPUT type="hidden" name="F_RecEmp5" value="<%=F_RecEmp(4)%>">

			<INPUT type="hidden" name="F_RecResults1" value="<%=F_RecResults(0)%>">
			<INPUT type="hidden" name="F_RecResults2" value="<%=F_RecResults(1)%>">
			<INPUT type="hidden" name="F_RecResults3" value="<%=F_RecResults(2)%>">
			<INPUT type="hidden" name="F_RecResults4" value="<%=F_RecResults(3)%>">
			<INPUT type="hidden" name="F_RecResults5" value="<%=F_RecResults(4)%>">

			<INPUT type="hidden" name="F_DelEmp1" value="<%=F_DelEmp(0)%>">
			<INPUT type="hidden" name="F_DelEmp2" value="<%=F_DelEmp(1)%>">
			<INPUT type="hidden" name="F_DelEmp3" value="<%=F_DelEmp(2)%>">
			<INPUT type="hidden" name="F_DelEmp4" value="<%=F_DelEmp(3)%>">
			<INPUT type="hidden" name="F_DelEmp5" value="<%=F_DelEmp(4)%>">

			<INPUT type="hidden" name="Email1" value="<%=Email1%>">
			<INPUT type="hidden" name="Email2" value="<%=Email2%>">
			<INPUT type="hidden" name="Email3" value="<%=Email3%>">
			<INPUT type="hidden" name="Email4" value="<%=Email4%>">
			<INPUT type="hidden" name="Email5" value="<%=Email5%>">

			<INPUT type="submit" value="ＯＫ" onClick="return GoEntry()">
			<INPUT type="submit" value="戻る" onClick="GoBack()">
		</TD>
	</TR>
</TABLE>
</FORM>
<% end if %>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
