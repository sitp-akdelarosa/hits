<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo292.asp				_/
'_/	Function	:事前空搬出指示書印刷画面		_/
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
  WriteLogH "b309", "空搬出指示書印刷", "02",""
  
  dim i,j,conInfo,Num
  Redim conInfo(5)
  j=0
  For i=0 To 4
    conInfo(j)=Array("","","","","","")
    conInfo(j)(0)=Request("ContSize"&i)
    conInfo(j)(1)=Request("ContType"&i)
    conInfo(j)(2)=Request("ContHeight"&i)
    conInfo(j)(3)=Request("Material"&i)
    conInfo(j)(4)=Request("PickPlace"&i)
    conInfo(j)(5)=Request("PickNum"&i)
    If conInfo(j)(0)="" AND conInfo(j)(1)="" AND conInfo(j)(2)="" AND conInfo(j)(3)="" AND conInfo(j)(4)="" AND conInfo(j)(5)="" Then
    Else
      j=j+1
    End If
  Next
  Num=j
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./styleprint.css">
<TITLE>指示書</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------空搬出指示書印刷画面--------------------------->
<CENTER><B class=titleB>空搬出指示書</B></CENTER>
<DIV style=text-align:right;>作成&nbsp;<%=Request("day")%></DIV><BR>
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD valign=top>指示者</TD><TD valign=top>＝<%=Request("SjManN")%></TD>
    <TD>（担当者：　　　　　　　　　　　　　　　）<BR>
        <%=Request("TelNo")%></TD></TR>
  <TR>
    <TD>作業者</TD><TD>＝<%=Request("WkManN")%></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>ブッキング番号　　・・・・・・</TD><TD><%=Request("BookNo")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=85% align=center>
  <TR><TD>項番</TD><TD>サイズ</TD><TD>タイプ</TD><TD>高さ</TD><TD>材質</TD><TD>ピック場所</TD><TD></TD><TD>本数</TD></TR>
<% For i=0 To Num-1%>
  <TR><TD><%=i+1%></TD>
      <TD><%=conInfo(i)(0)%>'</TD><TD><%=conInfo(i)(1)%></TD>
      <TD><%=conInfo(i)(2)%></TD><TD><%=conInfo(i)(3)%></TD>
      <TD><%=conInfo(i)(4)%></TD><TD>・・・</TD>
      <TD><%=conInfo(i)(5)%></TD></TR>
<% Next %>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>１．</TH>
    <TD nowrap><B>ブッキング情報&nbsp;</B></TD><TD></TD></TR>
  <TR>
    <TD>（船社）</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>（船名）</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>（仕向地）</TD><TD><%=Request("delivTo")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>２．</TH>
    <TD><B>バン詰め情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（バン詰め日時）</TD><TD><%=Request("vanDate")%></TD></TR>
  <TR>
    <TD valign=top nowrap>（バン詰め場所１）&nbsp;</TD><TD><%=Request("vanPlace1")%></TD></TR>
  <TR>
    <TD valign=top nowrap>（バン詰め場所２）</TD><TD><%=Request("vanPlace2")%></TD></TR>
  <TR>
    <TD>（品名）</TD><TD><%=Request("goodsName")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>３．</TH>
    <TD><B>搬入情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（搬入先ＣＹ）</TD><TD><%=Request("Terminal")%></TD></TR>
  <TR>
    <TD nowrap>（ＣＹカット日）&nbsp;</TD><TD><%=Request("CYCut")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>４．</TH>
    <TD><B>備考</B></TD><TD></TD></TR>
  <TR>
    <TD valign=top nowrap>（備考１）&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD valign=top>（備考２）</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<!-------------画面終わり--------------------------->
</BODY></HTML>
