<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo092.asp				_/
'_/	Function	:事前実搬出指示書印刷画面		_/
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
  WriteLogH "b109", "実搬出指示書印刷", "02",""
'コンテナデータ取得
  dim preConInfo,preNum,ConInfo,Num,i,j
  Get_Data preNum,preConInfo
  
  If Request("checkNum")="" OR Request("checkNum")=Null Then
    Num=1
    ReDim ConInfo(1)
    ConInfo(0)=preConInfo(0)
  Else
    dim strChecks,tmptarget,targetNo
    Num=Request("checkNum")
    strChecks=Request("checkeds")
    ReDim ConInfo(Num)
    tmptarget=Split(strChecks, ",", -1, 1)
    For i=0 To Num-1
      targetNo=Mid(tmptarget(i),3)
      ConInfo(i)=preConInfo(targetNo)
    Next
  End If
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
<!-------------実搬出指示書印刷画面--------------------------->
<CENTER><B class=titleB>実搬出指示書</B></CENTER>
<DIV style=text-align:right;>作成&nbsp;<%=Request("day")%></DIV><BR>
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD>作業番号</TD><TD>＝<%=Request("SakuNo")%></TD><TD></TD></TR>
  <TR>
    <TD valign=top>指示者</TD><TD valign=top>＝<%=Request("SjManN")%></TD>
    <TD>（担当者：　　　　　　　　　　　　　　　）<BR>
        <%=Request("TelNo")%></TD></TR>
  <TR>
    <TD>作業者</TD><TD>＝<%=Request("WkManN")%></TD>
    <TD>（ヘッドＩＤ＝<%=Request("HedId")%>）</TD></TR>
  <TR>
    <TD>指定方法</TD><TD>＝<%=Request("Way")%></TD><TD></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>１．</TH>
    <TD nowrap><B>コンテナ情報</B>&nbsp;</TD><TD></TD></TR>
  <TR>
    <TD>（船社）</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>（船名）</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>（品名）</TD><TD><%=Request("HinName")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>２．</TH>
    <TD><B>搬出情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（ＣＹ）</TD><TD><%=Request("Hfrom")%></TD></TR>
  <TR>
    <TD nowrap>（搬出予定日時）&nbsp;</TD><TD><%=Request("RDate")%></TD></TR>
  <TR>
    <TD valign=top>（納入先１）</TD><TD><%=Request("Nonyu1")%></TD></TR>
  <TR>
    <TD valign=top>（納入先２）</TD><TD><%=Request("Nonyu2")%></TD></TR>
  <TR>
    <TD>（納入日時分）</TD><TD><%=Request("NoDate")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>３．</TH>
    <TD><B>空コン返却情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（返却先）</TD><TD><%=Request("RPlace")%></TD></TR>
  <TR>
    <TD nowrap>（返却予定日数）&nbsp;</TD><TD><%=Request("Rnissu")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>４．</TH>
    <TD><B>備考</B></TD><TD></TD></TR>
  <TR>
    <TD valign=top nowrap>（備考１）&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD valign=top>（備考２）</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=<%=Num+3%> valign=top>５．</TH>
    <TD colspan=7><B>コンテナ番号</B></TD></TR>
  <TR><TD width=20></TD><TD>項番</TD><TD>&nbsp;コンテナ番号&nbsp;</TD><TD>&nbsp;サイズ&nbsp;</TD>
      <TD>&nbsp;タイプ&nbsp;</TD><TD>&nbsp;高さ&nbsp;</TD><TD>&nbsp;グロス&nbsp;</TD>
  <TR align=center><TD></TD>
    <TD>1</TD>
    <TD><%=ConInfo(0)(0)%></TD><TD><%=ConInfo(0)(1)%>'</TD><TD><%=ConInfo(0)(2)%></TD>
    <TD><%=ConInfo(0)(3)%></TD><TD><%=ConInfo(0)(4)%>kg</TD></TR>
<% For i=1 To Num-1 %>
  <TR align=center><TD></TD>
    <TD><%=i+1%></TD>
    <TD><%=ConInfo(i)(0)%></TD><TD><%=ConInfo(i)(1)%>'</TD><TD><%=ConInfo(i)(2)%></TD>
    <TD><%=ConInfo(i)(3)%></TD><TD><%=ConInfo(i)(4)%>kg</TD></TR>
<%Next%>
</TABLE><P>
<!-------------画面終わり--------------------------->
</BODY></HTML>
