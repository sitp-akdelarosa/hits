<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi020.asp				_/
'_/	Function	:事前実搬出一覧からの選択画面		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
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
  dim CONnum,Flag,i,j,Num
  CONnum = Request("CONnum")
  Flag   = Request("flag")

'コンテナ番号,会社コード取得
  dim param,CONnumA(),CMPcd(5)
  Num = Request("num")
  ReDim CONnumA(Num)
  For Each param In Request.Form
    If Left(param, 6) = "CONnum" Then
      If param <> "CONnum" Then
        i = Mid(param,7)
        CONnumA(i) = Request.Form(param)
      Else
        CONnumA(0)=CONnum
      End If
    ElseIf Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next

'コンテナ番号受渡しメソッド
Sub Set_CONnum
  For i = 1 to Num -1
    Response.Write "<INPUT type=hidden name='CONnum" & i & "' value='" & CONnumA(i) & "'>" & vbCrLf
  Next
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬出コンテナ選択</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//戻る
function GoBackT(){
  target=document.dmi020F;
  target.action="./dmi010.asp";
  len = target.elements.length;
  for (i=0; i<len-3; i++) target.elements[i].disabled = true;
  return true;
}
//登録
function GoEntry(){
  count=0;
  target=document.dmi020F;
  len = target.elements.length;
  for (i=0; i<len; i++){
    if(target.elements[i].checked)
    count++;
  }
  target.num.value=count;
  target.CONnum.disabled=false;
  target.action="./dmi021.asp";
  return true;
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------実搬出コンテナ選択画面--------------------------->
<FORM name="dmi020F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=3 width="100%">
  <TR>
    <TD colspan=2>
      <B>実搬出コンテナ選択</B></TD></TR>
  <TR hight=40>
    <TD>　</TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD><DIV class=bgb>コンテナ番号</DIV></TD></TR>
  <TR>
    <TD align=right>
      <INPUT type="checkbox" name="CONnum" value=<%=CONnumA(0)%> checked disabled></TD>
    <TD><DIV class=bgw><%=CONnumA(0)%></DIV></TD></TR>
<% For i=1 to Num-1%>
  <TR>
    <TD align=right><INPUT type="checkbox" name="CONnum<%=i%>" value=<%=CONnumA(i)%>></TD>
    <TD><DIV class=bgw><%=CONnumA(i)%></DIV></TD></TR>
<% Next%>
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name="BLnum" value="<%=Request("BLnum")%>">
       <INPUT type=hidden name="UpFlag"  value="<%=Request("UpFlag")%>">
       <INPUT type=hidden name=flag value="<%=Flag%>">
       <INPUT type=hidden name=num value="<%=Num%>">
       <INPUT type=hidden name="CONsize" value="<%=Request("CONsize")%>">
       <INPUT type=hidden name="CONtype" value="<%=Request("CONtype")%>">
       <INPUT type=hidden name="CONhite" value="<%=Request("CONhite")%>">
       <INPUT type=hidden name="CONtear" value="<%=Request("CONtear")%>">
       <INPUT type=hidden name="CMPcd0"  value="<%=CMPcd(0)%>">
       <INPUT type=hidden name="CMPcd1"  value="<%=CMPcd(1)%>">
       <INPUT type=hidden name="CMPcd2"  value="<%=CMPcd(2)%>">
       <INPUT type=hidden name="CMPcd3"  value="<%=CMPcd(3)%>">
       <INPUT type=hidden name="CMPcd4"  value="<%=CMPcd(4)%>">
       <INPUT type=hidden name="HFrom"   value="<%=Request("HFrom")%>">
<%'3th add   <INPUT type=hidden name="Comment3" value="<%=Comment3% >" > %>
       <INPUT type=hidden name="Rhou"     value="">
       <INPUT type=hidden name="shipFact" value="<%=Request("shipFact")%>" >
       <INPUT type=hidden name="shipName" value="<%=Request("shipName")%>" >
       <INPUT type=hidden name="HinName"  value="" >
       <INPUT type=hidden name="Nonyu1"   value="" >
       <INPUT type=hidden name="Nonyu2"   value="" >
       <INPUT type=hidden name="RPlace"   value="<%=Request("RPlace")%>" >
       <INPUT type=hidden name="Nomon"    value="">
       <INPUT type=hidden name="Noday"    value="">
       <INPUT type=hidden name="Nohou"    value="">
<%'3th add End %>
<%'       <INPUT type=submit value="キャンセル" onClick="return GoBackT()"> %>
       <INPUT type=submit value="ＯＫ" onClick="return GoEntry()">
    </TD></TR>
</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
