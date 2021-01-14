<%@LANGUAGE = VBScript%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi130.asp				_/
'_/	Function	:事前空搬入入力確認画面			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/07/29	備考欄追加	_/
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
  dim Mord,CONnum,CMPcd(5),HedId,Rmon,Rday
  dim param,i,j
  Mord   = Request("Mord")
  CONnum = Request("CONnum")
  For Each param In Request.Form
    If Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next

'表示文言生成
'3th  If Request("Rmon") = 0 Then 
'3th    Rmon = " "
'3th  Else
'3th    Rmon = Right("0"&Request("Rmon"),2)
'3th  End If
'3th
'3th  If Request("Rday") = 0 Then 
'3th    Rday = " "
'3th  Else
'3th    Rday = Right("0"&Request("Rday"),2)
'3th  End If

  dim ret
  If Mord=2 Then
    ret = true
  Else
  'DB接続
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS
  'ヘッドIDのチェック
    checkHdCd ObjConn, ObjRS, CMPcd, ret
  'DB接続解除
    DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
    on error goto 0
  End If
  dim tmpstr,tmpNo
  If Mord=0 Then 
    tmpNo="02"
  Else 
    tmpNo="14"
  End If
  If Request("UpFlag") <> 5 Then 
    tmpstr=CMpcd(Request("UpFlag"))&"/"
  Else
    tmpstr="/"
  End If
  tmpstr=tmpstr&Request("HedId")&"/"&Rmon&Rday&_
        "/"&Request("CONsize")&"/"&Request("CONtype")&"/"&Request("CONhite")&"/"&Request("CONsitu")&_
        "/"&Request("CONtear")&"/"&Request("MrSk")&"/"&Request("MaxW")
  If ret Then
    tmpstr=tmpstr&",入力内容の正誤:0(正しい)"
  Else
    tmpstr=tmpstr&",入力内容の正誤:1(誤り)"
  End If
'  WriteLogH "b202", "空搬入事前情報入力",tmpNo,tmpstr

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空搬入情報入力</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

function setParam(target){
  len = target.elements.length;
  for (i=0; i<len-2; i++) target.elements[i].readOnly = true;
  bgset(target);
}

//登録
function GoEntry(){
  target=document.dmi130F;
  target.action="./dmi140.asp";
  return true;
}
//戻る
function GoBackT(){
  target=document.dmi130F;
  target.action="./dmi120.asp";
  return true;
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi130F)">
<!-------------空搬入情報入力確認画面--------------------------->
<FORM name="dmi130F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2><B>空搬入情報入力確認</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>コンテナＮｏ．</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>"></TD></TR>
  <TR>
    <TD width=240><BR><DIV class=bgb>会社コード</DIV></TD>
    <TD>登録者<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5 maxlength=2>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>返却先</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>搬出予定日</DIV></TD>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=2>月
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=2>日
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>サイズ、タイプ、高さ、テアウェイト</DIV></TD>
    <TD><INPUT type=text name="CONsize" value="<%=Request("CONsize")%>" size=5>
        <INPUT type=text name="CONtype" value="<%=Request("CONtype")%>" size=5>
        <INPUT type=text name="CONhite" value="<%=Request("CONhite")%>" size=5>
        <INPUT type=text name="CONsitu" value="<%=Request("CONsitu")%>" size=5 style="display:none;">		<!-- 2016/10/24 H.Yoshikawa Upd (非表示とする) -->
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=5>kg
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>取扱船社</DIV></TD>
    <TD><INPUT type=text name="TrhkSen" value="<%=Request("TrhkSen")%>" size=27></TD></TR>
  <TR>
    <TD><DIV class=bgb>丸関</DIV></TD>
    <TD><INPUT type=text name="MrSk" value="<%=Request("MrSk")%>" size=5>
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>ＭＡＸ重量</DIV></TD>
    <TD><INPUT type=text name="MaxW" value="<%=Request("MaxW")%>" maxlength=6>kg</TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>備考</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73></TD></TR>
<%'C-002 ADD End %>

  <TR>
<!--  2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>登録担当者</DIV></TD>
  	<TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerSubName")%>" maxlength=16></TD>
<!--  2009/03/10 R.Shibuta Add-E -->
  </TR>
  
<% If Mord=1 AND Request("UpFlag")<>1 Then %>
  <TR>
    <TD colspan=2 align=center>
    <DIV class=bgw>指示元への回答　　　Yes　　　　　</DIV>
    </TD></TR>
    </TD></TR>
<% ElseIf Mord =2 Then %>
  <TR>
    <TD colspan=2 align=center>
    <DIV class=bgw>指示元への回答　　　No　　　　　</DIV>
    </TD></TR>
  <TR>
    <TD colspan=2 align=center>
       <DIV class=alert><B>＜注意＞</B>回答をNoで指定の場合は入力したデータは反映されません。</DIV>
    </TD></TR>
<% End If %>
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name=Mord value="<%=Mord%>" >
       <INPUT type=hidden name=UpFlag value="<%=Request("UpFlag")%>" >
       <INPUT type=hidden name=UpUser  value="<%=Request("UpUser")%>">
       <INPUT type=hidden name="compFlag"  value="<%=Request("compFlag")%>">
       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
<% If Not ret Then %>
       <P><DIV class=alert>
        指定された会社コードは存在しません。<BR>
       「戻る」ボタンを押下し、再入力してください。
       </DIV></P>
<% Else %>
       <INPUT type=submit value="ＯＫ" onClick="return GoEntry()">
<% End If %>
       <INPUT type=submit value="戻る" onClick="return GoBackT()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
