<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits											_/
'_/	FileName	:dmo312.asp										_/
'_/	Function	:事前実搬入情報取得、Booking件数をカウントし	_/
'_/				:		複数ならばdmi314.asp					_/
'_/				:		単数ならばShoriModeで制御				_/
'_/				:		"FLin"実搬入		dmi315.asp			_/
'_/				:		"EMoutInf"空搬出情報BookInfo()			_/
'_/				:		"EMoutUpd"空搬出登録dmi215.asp			_/
'_/	Date		:2006/03/06										_/
'_/	Code By		:SEIKO Electric.Co 松田勇人						_/
'_/	Modify		:												_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>

<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  
'データを取得
  dim CONnum,BookNo,User,ShipLine,ShoriMode,Mord
  dim WkNo													'2016/10/18 H.Yoshikawa Add
  BookNo = Trim(Request("BookNo"))
  CONnum = Trim(Request("CONnum"))
  ShoriMode = Trim(Request("ShoriMode"))
  ShipLine = Trim(Request("ShipLine"))
  Mord = Trim(Request("Mord"))
  User   = Session.Contents("userid")
  WkNo = gfTrim(Request("WkNo"))							'2016/10/18 H.Yoshikawa Add
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL, ErrerM
  ConnDBH ObjConn, ObjRS
  
'2016/10/18 H.Yoshikawa Add Start
'作業番号指定の場合は、CYVanInfoよりBookNoとShipLineを取得
  ret="0"
  if WkNo <> "" then
	StrSQL = "select BookNo, ShipLine "&_
			 "from CYVanInfo "&_
			 "where WkNo = '"& gfSQLEncode(WkNo) & "' "
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "1","b401","01","事前登録：作業番号存在チェック","101","SQL:<BR>"&StrSQL
	end if
	if ObjRS.eof then
		ret="3"
		ErrerM="指定した作業番号がシステムに登録されていません。<BR>入力の間違いがないか番号を確認してください。</P>"
	else
		BookNo = gfTrim(ObjRS("BookNo"))
		ShipLine = gfTrim(ObjRS("ShipLine"))
	end if
	ObjRS.Close
  end if
if ret = "0" then
'2016/10/18 H.Yoshikawa Add End

'ブッキング番号の存在チェック,ret=1:Booking0件(又はエラー),ret=0:Booking１件,ret=2:BookingN件
  dim dummy,ret
  ret="0"
  StrSQL = "select count(BOK.BookNo) as Num "&_
		   ",max( BOK.ShipLine) as ShipLine "&_
		   "from(select distinct BookNo,shipline from Booking) as BOK "&_
		   "where BOK.BookNo='"& BookNo & "' "
	if ShipLine<>"" then
		strsql=strsql & "and BOK.ShipLine='"& ShipLine & "' "
	end if
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b401","01","事前登録：重複チェック","101","SQL:<BR>"&StrSQL
  end if
  If Trim(ObjRS("Num")) = "0" Then
    ret="1"
    ErrerM="指定したブッキングNoがシステムに登録されていません。<BR>入力の間違いがないか番号を確認してください。</P>"
  ElseIf Trim(ObjRS("Num")) > "1" then
    ret="2"
    ErrerM="入力されたBookin番号は複数登録されています。</P>"
  else
    ShipLine = Trim(ObjRS("ShipLine"))
  End If
  ObjRS.Close
end if									'2016/10/18 H.Yoshikawa Add
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0
  
  If ret ="0"Then
    WriteLogH "b402", "実搬入事前情報入力", "01",BookNo&",入力内容の正誤:0(正しい)"
  elseif ret="2" then
    WriteLogH "b402", "実搬入事前情報入力", "01",BookNo&",入力内容の正誤:2(正しい、複数件)"
  Else
    WriteLogH "b402", "実搬入事前情報入力", "01",BookNo&",入力内容の正誤:1(誤り)"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>船会社選択</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){

<% If ret="0" Then %>
	<% If ShoriMode="FLin" Then %>
		window.resizeTo(850,690);
		target.action="./dmi315.asp";
		target.submit();
	<% elseIf ShoriMode="EMoutUpd" Then %>
		target.action="./dmi215.asp";
		target.submit();
	<% elseIf ShoriMode="EMoutInf" Then %>
		BookInfo(target);
	<% End If %>
<% elseIf ret="2" Then %>
  window.resizeTo(500,500);
  window.focus();
<% elseIf ret="1" Then %>
  window.resizeTo(500,500);
  window.focus();
<% End If %>
}

function GoNext(){
  target=document.dmi312F;
	document.dmi312F.action="./dmi314.asp";
    Num=LTrim(target.ShipLine.value );
    if(Num.length==0){
      alert("船社の頭文字を記入してください");
      target.ShipLine.focus();
      return;
    }

  if(!CheckEisu(target.ShipLine.value)){
    alert("頭文字に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
    target.ShipLine.focus();
    return;
  }
  chengeUpper(target);
  target.submit();
}
function GoBack(){
	<% If ShoriMode="FLin" Then %>
		target=document.dmi312F;
		target.action="./dmi310.asp";
	<% ElseIf ShoriMode="EMoutUpd" Then %>
      	window.open('dmi210.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
	<% Else %>
		target=document.dmi312F;
		window.close(); 
	<% End If %>
	target.submit();
}

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="setParam(document.dmi312F)">
<FORM name="dmi312F" method="POST">
<!-------------実搬入情報取得画面--------------------------->
<% If ret="0" Then %>

	<INPUT type=hidden name="BookNo" value="<%=BookNo%>">
	<INPUT type=hidden name="CONnum" value="<%=CONnum%>">
	<INPUT type=hidden name="ShipLine" value="<%=ShipLine%>">
	<INPUT type=hidden name="Mord" value="<%=Mord%>">
	<INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
	<INPUT type=hidden name="WkNo" value="<%=WkNo%>">						<!-- 2016/10/18 H.Yoshikawa Add -->


<% ElseIf ret="2" Then %>
<CENTER>
	<br><br><br>
  <DIV class=alert>
    <%= ErrerM%>
  </DIV>
	<table>
		<TD colspan="3" align="center">
			<br><br>
			<B>船社の頭文字を半角アルファベット1文字入力しＯＫを押してください。</B><BR>
			<br><br>
			<INPUT type=text  name="ShipLine" maxlength=1 size=3><BR>
			<INPUT type=hidden name="BookNo" value="<%=BookNo%>">
			<INPUT type=hidden name="CONnum" value="<%=CONnum%>">
			<INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
			<INPUT type=hidden name="Mord" value="<%=Mord%>">
			<br><br>
			<P><INPUT id=button1 type=button value="　戻る　" 
				name=button1 LANGUAGE=javascript onclick="GoBack()">
			<INPUT id=button1 type=button value="　ＯＫ　" 
				name=button1 LANGUAGE=javascript onclick="GoNext()"></P>
		</TD>
		</table>
</CENTER>

<% Else %>
<CENTER>
  <DIV class=alert>
    <%= ErrerM %>
  </DIV>
  <P><INPUT type=button value="閉じる" onClick="window.close()" id=button1 name=button1></P>
</CENTER>

<% End If %>

  <INPUT type=hidden name=DataNum value="<%=Request("Num")%>">
  <INPUT type=hidden name=SortFlag value="<%=Request("SortFlag")%>" >
  <INPUT type=hidden name=SortKye value="<%=Request("SortKye")%>" >
  <INPUT type=hidden name=CompF value="<%=Request("CompF")%>" >
  <INPUT type=hidden name=COMPcd0 value="<%=Request("COMPcd0")%>" >
  <INPUT type=hidden name=COMPcd1 value="<%=Request("COMPcd1")%>" >
  <INPUT type=hidden name=strWhere value="<%=Request("strWhere")%>">

  

</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>

