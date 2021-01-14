<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo291.asp				_/
'_/	Function	:事前空搬出指示書印刷調整画面		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:								_/
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
'ログ出力
  WriteLogH "b309", "空搬出指示書印刷", "01",""

'サーバ日時の取得
  dim DayTime,day
  getDayTime DayTime
  day = DayTime(0) & "年" & DayTime(1) & "月" & DayTime(2) & "日" 

'前画面からのデータ取得
  dim vanDate,vanTime,YY,i
  dim COMPcd1,vanMon,vanDay
  COMPcd1 = Request("COMPcd1")

'日の整形
  vanMon =Right("00" & Request("vanMon"),2)
  vanDay =Right("00" & Request("vanDay"),2)
  If Request("vanHou")= "" Then
    vanTime=""
  Else
    vanTime=Right("00" & Request("vanHou"),2) & "時" & Right("00" & Request("vanMin"),2) & "分"
  End IF
  '日の年度を決定
  If DayTime(1) > vanMon Then	'来年
    YY = DayTime(0) +1
  ElseIf DayTime(1) = vanMon AND DayTime(2) > vanDay Then
    YY = DayTime(0) +1
  Else
    YY = DayTime(0)
  End If
  If vanMon = "00" Or vanDay = "00" Then
    vanDate= ""
  Else
    vanDate= YY &"年"& vanMon &"月"& vanDay &"日　"& vanTime
  End If

  
'セッションからユーザ名称を取得
  Dim SjManN
  SjManN = Session.Contents("LinUN")

'DBからのデータ取得
  'エラートラップ開始
  on error resume next
  'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  '作業者名称取得
  Dim WkManN
  If Trim(COMPcd1)="" OR COMPcd1=Null Then
    WkManN=SjManN
  Else
    StrSQL = "Select FullName From mUsers Where HeadCompanyCode='" & COMPcd1 &"'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b309","01","空搬出指示書印刷調整・作業者名取得","102","SQL:<BR>"&strSQL
    end if
    WkManN= Trim(ObjRS("FullName"))
    ObjRS.close
  End If
'指示者電話番号取得
  dim USER,TelNo
  USER       = Session.Contents("userid")
  StrSQL = "select TelNo from mUsers where UserCode='" & USER &"'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b309","01","空搬出指示書印刷調整・指示者電話番号取得","102","SQL:<BR>"&strSQL
  end if
  TelNo = Trim(ObjRS("TelNo"))
  ObjRS.close
  If TelNo<>"" Then
    TelNo="（電話番号："&TelNo&"）"
  End If

  'DB接続解除
  DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>指示書印刷調整</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
//指示書印刷画面へ
function GoNext(){
  target=document.dmo291F;
  newWin = window.open("", "Print2", "width=950,height=850,left=10,top=10,resizable=yes,scrollbars=yes,menubar=yes,top=0");
  target.target="Print2";
  target.submit();
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmo291F.WkManN.focus();
}
//2008-01-31 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onload="finit();">
<!-------------空搬出指示書印刷調整画面--------------------------->
<FORM name="dmo291F" method="POST" action="./dmo292.asp";>
<CENTER><B class=titleB>空搬出指示書</B></CENTER>
<DIV style=text-align:right;>作成&nbsp;<%=day%></DIV>
<INPUT type=hidden name="day" value="<%=day%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR>
    <TD valign=top>指示者</TD><TD valign=top>＝<%=SjManN%></TD>
    <TD>（担当者：　　　　　　　　　　　　　　　）<BR>
        <%=TelNo%></TD></TR>
  <TR>
    <TD>作業者</TD><TD>＝<INPUT type=text name="WkManN" value="<%=WkManN%>"></TD><TD></TD></TR>
  <TR>
    <TD colspan=2>ブッキング番号　　・・・・・・</TD><TD><%=Request("BookNo")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="SjManN" value="<%=SjManN%>">
<INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>">
<INPUT type=hidden name="TelNo" value="<%=TelNo%>">

<TABLE border=0 cellPadding=0 cellSpacing=0 width=85% align=center>
  <TR><TD></TD><TD>サイズ</TD><TD>タイプ</TD><TD>高さ</TD><TD>材質</TD><TD>ピック場所</TD><TD></TD><TD>本数</TD></TR>
<% For i=0 To 4%>
  <TR><TD>(<%=i+1%>)</TD>
      <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4 maxlength=2>'</TD>
      <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4 maxlength=2></TD>
      <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4 maxlength=2></TD>
      <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4 maxlength=1></TD>
      <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 maxlength=20></TD>
      <TD>・・・</TD>
      <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4   maxlength=3></TD></TR>
<% Next %>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>１．</TH>
    <TD><B>ブッキング情報&nbsp;</B></TD><TD></TD></TR>
  <TR>
    <TD>（船社）</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>（船名）</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>（仕向地）</TD><TD><%=Request("delivTo")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="shipFact" value="<%=Request("shipFact")%>">
<INPUT type=hidden name="shipName" value="<%=Request("shipName")%>">
<INPUT type=hidden name="delivTo"  value="<%=Request("delivTo")%>">

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>２．</TH>
    <TD><B>バン詰め情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（バン詰め日時）</TD><TD><%=vanDate%></TD></TR>
  <TR>
    <TD>（バン詰め場所１）&nbsp;</TD><TD><%=Request("vanPlace1")%></TD></TR>
  <TR>
    <TD>（バン詰め場所２）</TD><TD><%=Request("vanPlace2")%></TD></TR>
  <TR>
    <TD>（品名）</TD><TD><%=Request("goodsName")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="vanDate"  value="<%=vanDate%>">
  <INPUT type=hidden name="vanPlace1" value="<%=Request("VanPlace1")%>">
  <INPUT type=hidden name="vanPlace2" value="<%=Request("VanPlace2")%>">
  <INPUT type=hidden name="goodsName" value="<%=Request("GoodsName")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>３．</TH>
    <TD><B>搬入情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（搬入先ＣＹ）</TD><TD><%=Request("Terminal")%></TD></TR>
  <TR>
    <TD>（ＣＹカット日）&nbsp;</TD><TD><%=Request("CYCut")%></TD></TR>
</TABLE><P>
  <INPUT type=hidden name="Terminal"  value="<%=Request("Terminal")%>">
  <INPUT type=hidden name="CYCut"    value="<%=Request("CYCut")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>４．</TH>
    <TD><B>備考</B></TD><TD></TD></TR>
  <TR>
    <TD>（備考１）&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD>（備考２）</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="Comment1"  value="<%=Request("Comment1")%>">
<INPUT type=hidden name="Comment2"  value="<%=Request("Comment2")%>">
<CENTER>
  <INPUT type=button value="ＯＫ" onClick="GoNext()">
  <INPUT type=button value="閉じる" onClick="window.close()">
</CENTER>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
