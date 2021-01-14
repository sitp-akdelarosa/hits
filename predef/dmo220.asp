<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo220.asp				_/
'_/	Function	:事前空搬出入力表示画面			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/06	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次全面改修	_/
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
  WriteLogH "b302", "空搬出事前情報入力","11",""

'データ所得
  dim COMPcd0,ret,compF,i
  COMPcd0= Request("COMPcd0")
  compF  = Request("compF")
  
  Const RowNum = 10					'2017/05/09 H.Yoshikawa Add

'更新モードフラグ設定
  ret=true
  If compF<>0 AND COMPcd0 <> UCase(Session.Contents("userid")) Then
    ret=false
  End If
  
  dim WkOutFlag, OutStyle							'2016/08/25 H.Yoshikawa Add

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空バンピック情報表示</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  // 2016/10/24 H.Yoshikawa Upd Start
  //window.resizeTo(550,680);
  window.moveTo(120,20);
  window.resizeTo(1366,768);			// 2017/05/09 H.Yoshikawa Upd(770⇒820) // edited by AK.DELAROSA 2021-01-14
  // 2016/10/24 H.Yoshikawa Upd End
  window.focus();
  bgset(target);
}
//更新画面へ
function GoReEntry(){
  target=document.dmo220F;
  target.action="./dmi220.asp";
  target.submit();
}
//ブッキング情報
function GoBookI(){
  target=document.dmo220F
  BookInfo(target);
}
//指示書印刷調整画面へ
function GoSijiPrint(){
  target=document.dmo220F;
  target.action="./dmo291.asp";
//  newWin = window.open("", "Print", "width=500,height=700,left=30,top=10,resizable=yes,scrollbars=yes,top=0");
//  target.target="Print";
  target.submit();
//  target.target="_self";
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo220F)">
<!-------------空搬出情報表示確認画面--------------------------->
<FORM name="dmo220F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>空バンピック情報入力(表示モード)</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>ブッキングＮｏ．</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>船社</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>船名</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readOnly size=40>
    	<INPUT type=hidden name="VslCode" value="<%=Request("VslCode")%>">							<!-- 2016/08/22 H.Yoshikawa Add -->
    </TD></TR>
  <TR>
  	<!-- 2016/08/22 H.Yoshikawa Upd Start -->
    <!--<TD><DIV class=bgb>仕向地</DIV></TD>
    <TD><INPUT type=text name="delivTo" value="<%=Request("delivTo")%>" readOnly size=40></TD></TR> -->
    <TD><DIV class=bgb>Voyage</DIV></TD>
    <TD><INPUT type=hidden name="delivTo" value="<%=Request("delivTo")%>">
    	<INPUT type=text name="ExVoyage" value="<%=Request("ExVoyage")%>" readOnly size=12>			<!-- 2016/10/17 H.Yoshikawa Add -->
     	<INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>" >							<!-- 2016/10/17 H.Yoshikawa Upd(text→hidden) -->
   </TD></TR>
  	<!-- 2016/08/22 H.Yoshikawa Upd End -->
  <TR>
    <TD><DIV class=bgb>会社コード(陸運)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=Request("COMPcd1")%>" size=5  readOnly>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>属性と本数</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=1 cellSpacing=0 width="90%" align=center>
    <!-- 2016/08/16 H.Yoshikawa Upd Start -->
    <!-- <TR><TD></TD><TD>サイズ</TD><TD>タイプ</TD><TD>高さ</TD><TD>材質</TD><TD>ピック場所</TD><TD></TD><TD>本数</TD></TR> -->
    <TR>
    	<TD></TD>
    	<TD>サイズ</TD>
    	<TD>タイプ</TD>
    	<TD>高さ</TD>
    	<TD>設定温度</TD>
    	<TD>プレクール</TD>
    	<TD>ベンチレーション</TD>
    	<TD>ピック予定日時(時間はﾌﾟﾚｸｰﾙ時のみ必須)</TD>
    	<TD>　本数</TD>
    	<TD>搬出可否</TD>
    	<TD>ピックアップ場所</TD>
    	<TD>行削除</TD>									<!-- 2017/05/10 H.Yoshikwawa Add -->
    </TR>
    <!-- 2016/08/16 H.Yoshikawa Upd End -->
<% For i=0 To RowNum - 1 %>		<!-- 2017/05/09 H.Yoshikawa Upd(4⇒RowNum-1) -->
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4  readOnly></TD>
      <!-- 2016/08/22 H.Yoshikawa Upd Start
          <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4  readOnly></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 readOnly></TD>
          <TD>・・・</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4  readOnly></TD></TR> -->
          <TD><INPUT type=text name="SetTemp<%=i%>"  value="<%=Request("SetTemp"&i)%>" size=8 readOnly>℃</TD>
          <TD>
			<select disabled>
				<option value="0"></option>
				<option value="1" <% if gfTrim(Request("Pcool"&i)) = "1" then %>selected<% end if %> >有</option>
				<option value="2" <% if gfTrim(Request("Pcool"&i)) = "2" then %>selected<% end if %> >無</option>	<!-- 2017/08/25 H.Yoshikawa Add -->
			</select>
          <INPUT type=hidden name="Pcool<%=i%>"  value="<%=Request("Pcool"&i)%>"></TD>
          <TD><INPUT type=text name="Ventilation<%=i%>"  value="<%=Request("Ventilation"&i)%>" size=5 readOnly>%（開口）</TD>
          <TD>
              <INPUT type=text name="PickDate<%=i%>"  value="<%=Request("PickDate"&i)%>" size=15 readOnly>
              <INPUT type=text name="PickHour<%=i%>"  value="<%=Request("PickHour"&i)%>" size=4 readOnly>時
              <INPUT type=text name="PickMinute<%=i%>"  value="<%=Request("PickMinute"&i)%>" size=4 readOnly>分
          </TD>
          <TD>…<INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4 readOnly>
          <% OutStyle = ""
             select case Trim(Request("OutFlag"&i))
               case "0"
                 WkOutFlag = "確認中"
               case "1"
                 WkOutFlag = "可"
               case "9"
                 WkOutFlag = "不可"
                 OutStyle = "color:red;"
               case else
                 WkOutFlag = ""
             end select
          %>
          </TD>
          <TD style="<%=OutStyle%>"><INPUT type=hidden name="OutFlag<%=i%>"  value="<%=Request("OutFlag"&i)%>" ><%=WkOutFlag %></TD>
          <TD><INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"><%=gfHTMLEncode(Request("PickPlace"&i))%>
              <INPUT type=hidden name="Terminal<%=i%>"  value="<%=Request("Terminal"&i)%>">				
          </TD>
          <% '2017/05/10 H.Yoshikawa Upd Start %>
          <TD><INPUT type=checkbox value="1" disabled <% if Request("DelFlag"&i) = "1" then%> checked <% end if %>>
              <INPUT type=hidden name="DelFlag<%=i%>" value="<%=Request("DelFlag"&i)%>">
          </TD>
		  <% '2017/05/10 H.Yoshikawa Upd End %>
              <INPUT type=hidden name="UpdFlag<%=i%>"    <% if gfTrim(Request("ContSize"&i)) = "" then %>value="0" <% else %> value="0" <% end if %>>
              
	  </TR>
      <!-- 2016/08/22 H.Yoshikawa Upd End -->
		<% '2016/10/27 H.Yoshikawa Upd Start %>
		<INPUT type=hidden name="Bef_ContSize<%=i%>"    value="<%=Request("Bef_ContSize"&i)%>">
		<INPUT type=hidden name="Bef_ContType<%=i%>"    value="<%=Request("Bef_ContType"&i)%>">
		<INPUT type=hidden name="Bef_ContHeight<%=i%>"  value="<%=Request("Bef_ContHeight"&i)%>">
		<INPUT type=hidden name="Bef_SetTemp<%=i%>"     value="<%=Request("Bef_SetTemp"&i)%>">
		<INPUT type=hidden name="Bef_Pcool<%=i%>"       value="<%=Request("Bef_Pcool"&i)%>">
		<INPUT type=hidden name="Bef_Ventilation<%=i%>" value="<%=Request("Bef_Ventilation"&i)%>">
		<INPUT type=hidden name="Bef_PickDate<%=i%>"    value="<%=Request("Bef_PickDate"&i)%>">
		<INPUT type=hidden name="Bef_PickHour<%=i%>"    value="<%=Request("Bef_PickHour"&i)%>">
		<INPUT type=hidden name="Bef_PickMinute<%=i%>"  value="<%=Request("Bef_PickMinute"&i)%>">
		<INPUT type=hidden name="Bef_PickNum<%=i%>"     value="<%=Request("Bef_PickNum"&i)%>">
		<INPUT type=hidden name="Bef_OutFlag<%=i%>"     value="<%=Request("Bef_OutFlag"&i)%>">
		<INPUT type=hidden name="Bef_PickPlace<%=i%>"   value="<%=Request("Bef_PickPlace"&i)%>">
		<INPUT type=hidden name="Bef_Terminal<%=i%>"    value="<%=Request("Bef_Terminal"&i)%>">
		<% '2016/10/27 H.Yoshikawa Upd End %>
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め日時</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3  readOnly>月
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3  readOnly>日
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3  readOnly>時
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3  readOnly>分
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め場所１</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め場所２</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>品名</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>搬入先ＣＹ．ＣＹカット日</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考１</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考２</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73  readOnly></TD></TR>
    
  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
   <TD><DIV class=bgy>登録担当者</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" value="<%=Request("TruckerSubName")%>" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>
<!-- 2016/08/22 H.Yoshikawa Add Start -->
  <TR>
  	<TD><DIV class=bgy>電話番号</DIV></TD>
 	<TD><INPUT type=text name="Tel" value="<%=Request("Tel")%>"  readonly></TD></TR>
  <TR>
  	<TD><DIV class=bgy>メールアドレス</DIV></TD>
 	<TD><INPUT type=text name="Mail" value="<%=Request("Mail")%>" readonly size=60>
 		<INPUT type=checkbox value="1" <% if Request("MailFlag") = "1" then %>checked <% end if %> disabled>
 		搬出可否状態変更時にメールを受け取る
 		<INPUT type=hidden name="MailFlag" value="<%=Request("MailFlag")%>">
 	</TD></TR>
<!-- 2016/08/22 H.Yoshikawa Add End -->
  
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name=Mord value="<%=Request("Mord")%>" >
       <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
       <INPUT type=hidden name="TFlag" value="<%=Request("TFlag")%>">
<%'Add-s 2006/03/06 h.matsuda%>
       <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" >
	   <INPUT type=hidden name="ShoriMode" value="EMoutInf">
<%'Add-e 2006/03/06 h.matsuda%>
<%' If COMPcd0 = UCase(Session.Contents("userid")) Then  '''Del 20040301%>
       <INPUT type=button value="指示書印刷" onClick="GoSijiPrint()">
<%' End If '''Del 20040301%>
<% If ret Then %>
       <INPUT type=hidden name="compFlag" value="<%=compF%>">
       <INPUT type=submit value="更新モード" onClick="GoReEntry()">
<% End If %>
       <INPUT type=submit value="閉じる" onClick="window.close()">
       <P>
       <INPUT type=button value="ブッキング情報" onClick="GoBookI()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
