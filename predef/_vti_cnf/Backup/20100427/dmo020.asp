<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo020.asp				_/
'_/	Function	:事前実搬出入力画面(表示)		_/
'_/	Date		:2003/05/27				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/07/29	備考欄追加	_/
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
  dim SakuNo,Flag,UpFlag,Num,CONnumA(),CMPcd(5),Rmon,Rday
  dim param,i,j,Way
  Way   =Array("","指定あり","指定なし","一覧から選択","ＢＬ番号")
  SakuNo= Request("SakuNo")
  Flag= Request("flag")
  WriteLogH "b10"&(2+Flag), "実搬出事前情報入力("&Way(Flag)&")", "11",""
  Num = Request("num")
  UpFlag=Request("UpFlag")
  ReDim CONnumA(Num)
  i=0
  For Each param In Request.Form
    If Left(param, 6) = "CONnum" Then
      CONnumA(i) = Request.Form(param)
      i=i+1
    ElseIf Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next

'表示文言生成
'3th del  If Request("Rmon") = "" Then 
'3th del    Rmon = "-"
'3th del  Else
'3th del    Rmon = Request("Rmon")
'3th del  End If
'3th del
'3th del  If Request("Rday") = "" Then 
'3th del    Rday = "-"
'3th del  Else
'3th del    Rday = Request("Rday")
'3th del  End If

'コンテナ番号受渡しメソッド
Sub Set_CONnum
  For i = 1 to Num -1
    Response.Write "       <INPUT type=hidden name='CONnum" & i & "' value='" & CONnumA(i) & "'>" & vbCrLf
  Next
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬出情報入力(表示)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
var w=600;
var h=870;
if(screen.width){
        l=(screen.width-w)/2;
    }
    if(screen.availWidth){
        l=(screen.availWidth-w)/2;
    }
    if(screen.height){
        t=(screen.height-h)/2;
    }
    if(screen.availHeight){
        t=(screen.availHeight-h)/2;
    }
//window.resizeTo(600,770);
window.resizeTo(w,h);
window.scrollTo(w, h);

window.focus();

function setParam(target){
  for (i=0; i<29; i++) target.elements[i].readOnly = true;
  bgset(target);
}

//コンテナ詳細画面
function GoConInfo(){
  target=document.dmo020F;
  ConInfo(target,<%=Flag%>,0);
  return false;
}
//更新画面へ
function GoReEntry(){
  target=document.dmo020F;
  target.action="./dmi021.asp";
  return true;
}
//指示書印刷調整画面へ
function GoSijiPrint(){
  window.resizeTo(500,700);
  target=document.dmo020F;
  target.action="./dmo091.asp";
//  newWin = window.open("", "Print", "width=500,height=700,left=30,top=10,resizable=yes,scrollbars=yes,top=0");
//  target.target="Print";
  target.submit();
//  target.target="_self";
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------実搬出情報入力(表示)画面--------------------------->
<FORM name="dmo020F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD><B>実搬出情報入力(表示モード)</B></TD>
    <TD><TABLE border=1 cellPadding=3 cellSpacing=0 align="right">
          <TR bgcolor="#f0f0f0"><TD>作業番号</TD><TD><%=SakuNo%></TD></TR>
        </TABLE>
    </TD></TR>
  <TR>
<% If Flag=4 Then %>
    <TD><DIV class=bgb>ＢＬＮｏ．</DIV></TD>
    <TD><INPUT type=text name="BLnum" value="<%=Request("BLnum")%>">　　<%=Way(Flag)%>
    <INPUT type=hidden name="CONnum" value="<%=CONnumA(0)%>"></TD></TR>
<% Else %>
    <TD><DIV class=bgb>コンテナＮｏ．</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnumA(0)%>">　　<%=Way(Flag)%></TD></TR>
    <INPUT type=hidden name="BLnum" value="<%=Request("BLnum")%>"></TD></TR>
<% End If %>
  <TR>
    <TD width=180>
        <DIV class=bgb>サイズ、タイプ、高さ、グロス</DIV></TD>
    <TD><INPUT type=text name="CONsize" value="<%=Request("CONsize")%>" size=5>
        <INPUT type=text name="CONtype" value="<%=Request("CONtype")%>" size=5>
        <INPUT type=text name="CONhite" value="<%=Request("CONhite")%>" size=5>
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=5>kg
    </TD></TR>
<%'3th追加 Start%>
  <TR>
    <TD><DIV class=bgb>船社、船名</DIV></TD>
    <TD><INPUT type=text name="Shipfact" value="<%=Request("shipFact")%>" size=20>
        <INPUT type=text name="ShipName" value="<%=Request("shipName")%>" size=20>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>品名</DIV></TD>
    <TD><INPUT type=text name="HinName" value="<%=Request("HinName")%>" size=40 maxlength=20>
    </TD></TR>
<%'3th追加 End%>
  <TR>
    <TD><BR><DIV class=bgb>会社コード</DIV></TD>
    <TD>登録者<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5 maxlength=2>
    </TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD Align=right>指示元担当者</TD>
    <TD>
        <INPUT type=text name="SubName" readonly = "readonly" value="<%=Request("TruckerSubName")%>" maxlength=16>
    </TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>ＣＹ</DIV></TD>
    <TD><INPUT type=text name="HFrom" value="<%=Request("Hfrom")%>"></TD></TR>
    <TD><DIV class=bgb>搬出予定日</DIV></TD>
<%'chage 3th    <TD><select name="Rmon" onchange="check_date('<%=DayTime(0)% >','<%=DayTime(1)% >',dmi021F.Rmon,dmi021F.Rday)">
'        </select>月<select name="Rday"></select>日 %>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=3 maxlength=2>月
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=3 maxlength=2>日
        <INPUT type=text name="Rhou" value="<%=Request("Rhou")%>" size=3 maxlength=2>時
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>搬出先</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" size=40></TD></TR>
<%'3th追加 Start%>
  <TR>
    <TD><DIV class=bgb>納入先１</DIV></TD>
    <TD><INPUT type=text name="Nonyu1" value="<%=Request("Nonyu1")%>" size=73>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>納入先２</DIV></TD>
    <TD><INPUT type=text name="Nonyu2" value="<%=Request("Nonyu2")%>" size=73>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>納入日時</DIV></TD>
    <TD><INPUT type=text name="Nomon" value="<%=Request("Nomon")%>" size=3 maxlength=2>月
        <INPUT type=text name="Noday" value="<%=Request("Noday")%>" size=3 maxlength=2>日
        <INPUT type=text name="Nohou" value="<%=Request("Nohou")%>" size=3 maxlength=2>時
		<!-- 2008/01/31 Add S G.Ariola -->
		<INPUT type=text name="Nomin" value="<%=Request("Nomin")%>" size=3 maxlength=2>分
		<!-- 2008/01/31 Add E G.Ariola -->
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>空コン返却先</DIV></TD>
    <TD><INPUT type=text name="RPlace" value="<%=Request("RPlace")%>" size=30>
    </TD></TR>
<%'3th追加 End%>
  <TR>
    <TD><DIV class=bgb>返却予定日数（フリータイム）</DIV></TD>
    <TD><INPUT type=text name="Rnissu" value="<%=Request("Rnissu")%>">
    </TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>備考１</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考２</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73></TD></TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
  <TR>
   <TD><DIV class=bgy>登録担当者</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>

<%'Del 3th  <TR>
'    <TD><DIV class=bgb>備考３</DIV></TD>
'    <TD><INPUT type=text name="Comment3" value="<%=Request("Comment3")% >" size=13 maxlength=10></TD></TR>%>
<%'C-002 ADD End %>
  <TR>
    <TD colspan=2 align=center>
       <DIV class=alert><B>＜注意＞</B>船社によってはゲートで認証IDの入力が必要になります</DIV>
    </TD></TR>
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name="UpUser" value="<%=Request("UpUser")%>" >
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
 <!-- 2009/08/04 Tanaka Add-S -->
       <INPUT type=hidden name="TruckerName" value="<%=Request("TruckerName")%>" >
 <!-- 2009/08/04 Tanaka Add-E -->
<% If Num > 1 Then call Set_CONnum End If%>
       <INPUT type=button value="指示書印刷" onClick="GoSijiPrint()">
       <INPUT type=hidden name="SakuNo" value="<%=SakuNo%>">
       <INPUT type=hidden name=flag value="<%=Flag%>" >
       <INPUT type=hidden name=num value="<%=Num%>" >
<%'20030909 IF Request("compFlag") AND (UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1) Then %>

       <!--INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>"-->
<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
' IF UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1 Then 
  IF UCase(Session.Contents("userid"))=CMPcd(0) Or (Request("compFlag") AND Request("TruckerFlag")<>1) Then %>
       <INPUT type=submit value="更新モード" onClick="return GoReEntry()">
       <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
       <INPUT type=hidden name="WkCNo"    value="<%=Request("WkCNo")%>">
       <INPUT type=hidden name="TruckerFlag" value="<%=Request("TruckerFlag")%>">
<%End IF%>
       <INPUT type=submit value="閉じる" onClick="window.close()">
       <P>
       <INPUT type=submit value="コンテナ情報" onClick="return GoConInfo()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
<SCRIPT language=JavaScript>
setParam(document.dmo020F);
</SCRIPT>
</BODY></HTML>
