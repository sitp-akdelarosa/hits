<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi310.asp				_/
'_/	Function	:事前実搬入番号入力画面		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%><% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b402", "実搬入事前情報入力","00",""
  
  Dim ActionType 
  ActionType = Trim(Request.QueryString("ActionType"))
  'Y.TAKAKUWA Add-S 2015-03-13
  Dim CheckDigit
  CheckDigit = Trim(Request.QueryString("CheckDigit"))
  'Y.TAKAKUWA Add-E 2015-03-13
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE><%If ActionType <> "M" Then %>事前登録・搬入票作成<%End If%></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>

<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT src="./JS/CommonSub.js"></SCRIPT>

<SCRIPT language=JavaScript>
<!--
<% If ActionType <> "M" Then %>
window.resizeTo(400,400);
window.focus();
<%End If%>

function GoNext(){
  strA    = new Array("ブッキング番号","コンテナ番号","作業番号");		// 2016/10/18 H.Yoshikawa Upd（作業番号追加）
  target=document.dmi310F;

// 2016/10/18 H.Yoshikawa Add Start
  if(Rtrim(target.BookNo.value, " ")=="" && Rtrim(target.WkNo.value, " ")==""
   || Rtrim(target.BookNo.value, " ")!="" && Rtrim(target.WkNo.value, " ")!=""){
      alert("ブッキング番号、または、作業番号いずれかを記入してください");
      target.BookNo.focus();
      return;
  }
// 2016/10/18 H.Yoshikawa Add End

  targetA    = new Array();
  targetA[0] = target.BookNo;
  targetA[1] = target.CONnum;
  targetA[2] = target.WkNo;												// 2016/10/18 H.Yoshikawa Add
  for(k=0;k<2;k++){
   if(k==1){															// 2016/10/18 H.Yoshikawa Add
    Num=LTrim(targetA[k].value);
    if(Num.length==0){
      alert(strA[k]+"を記入してください");
      targetA[k].focus();
      return;
    }
   }																	// 2016/10/18 H.Yoshikawa Add
    if(k==0){
      if(!CheckEisu(targetA[k].value)){
        alert(strA[k]+"に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
        targetA[k].focus();
        return;
      }
    }else{
      if(!CheckEisu2(targetA[k].value)){
        alert(strA[k]+"に半角英数字以外の文字を記入しないでください");
        targetA[k].focus();
        return;
      }
    }
  }
  
  //Y.TAKAKUWA Add-S 2015-03-13
  var chkDigit;
  chkDigit = gfJDigitCheck(targetA[1]);
  //Y.TAKAKUWA Add-E 2015-03-13
  //Y.TAKAKUWA Upd-S 2015-03-13
  //alert(chkDigit);
  //var retValue = showModalDialog ("dmi310.asp?ActionType=M", window, "dialogWidth:330px; dialogHeight:80px; center:1; scroll: no; dialogTop:300px; ");
  var retValue;
  if(chkDigit == 0) 
  {
    //2016/11/17 H.Yoshikawa Upd Start
    //retValue = showModalDialog ("dmi310.asp?CheckDigit=" + chkDigit + "&ActionType=M", window, "dialogWidth:370px; dialogHeight:80px; center:1; scroll: no; dialogTop:300px; ");
     chengeUpper(target);
     target.submit();               
    //2016/11/17 H.Yoshikawa Upd End
  }
  else
  {
	retValue = showModalDialog ("dmi310.asp?CheckDigit=" + chkDigit + "&ActionType=M", window, "dialogWidth:450px; dialogHeight:100px; center:1; scroll: no; dialogTop:300px; ");
  }
  //Y.TAKAKUWA Upd-E 2015-03-13
  if (retValue) {
     chengeUpper(target);
     target.submit();               
  }

}
//2008-01-31 Add-S M.Marquez
function finit(){  
    <% If ActionType <> "M" Then %>
    document.dmi310F.BookNo.focus();
    <%End If%>
}
//2008-01-31 Add-E M.Marquez

function fStop()
{
  returnValue = false;
  window.close();
}
function fSend()
{
  returnValue = true;
  window.close();
}
//Y.TAKAKUWA Add-S 2015-03-13
//**************************************************
//  機能   : コンテナ番号のディジットチェックを行う
//
//  引数   : sContNo           As String     - [I] コンテナ番号
//
//  戻り値 ：チェック結果
//             0 - 正常
//             1 - 計算不能コンテナ
//             9 - チェックディジットエラー
//            -1 - 例外エラー
//**************************************************
function gfJDigitCheck(sContNo){

    var LsChar1;    //１文字エリア
    var LsChar4;    //４文字エリア
    var LsChar6;    //６文字エリア
    var LsWkContNo; //コンテナＮＯ（大文字）
    var LiIdx1;     //添字
    var LiIdx2;     //添字
    var LiIdx3;     //添字
    var LiData1;    //計算エリア
    var LiAmari;    //計算エリア
    var LiLen;      //長さ
    var LlData = 0;     //計算エリア
    var LsDigit;    
    var snum;    
    
    LiIdx2 = 0;
    LiIdx3 = 0;
    LsWkContNo = sContNo.value.toUpperCase();
       
    LiLen = sContNo.value.length;
    
    //入力ありなしチェック
    if(LiLen==0){
        return(1);
    }
     
    for (LiIdx1 = 1; LiIdx1 <= LiLen; LiIdx1++) {
        //ラテン文字の変換コード使用
        //65: "A" ～ 90: "Z"  
        snum = LsWkContNo.charCodeAt(LiIdx1);
        if (snum >= 65 && snum <= 90){
            LiIdx2 = LiIdx2 + 1;
        }else{ 
            break;
        }
    }
    
    //ＰＲＥＦＩＸの妥当性チェック
    if(LiIdx2 == 0 || LiIdx2 < 3 ){
        return(1);
    }
 
    LsChar4 = LsWkContNo.substring(0, LiIdx2);

    //番号部６桁チェック 
    //48: "0" ～ 57: "9"  
    for (LiIdx1 = LiIdx2 + 1; LiIdx1 <= 12; LiIdx1++) {
        //ラテン文字の変換コード使用
        //48: "0" ～ 57: "9"  
        snum = LsWkContNo.charCodeAt(LiIdx1);
        if (snum >= 48 && snum <= 57){
            LiIdx3 = LiIdx3 + 1;
        }else{ 
            break;
        }
    }
    //番号部６～７桁以外エラー
    if(LiIdx3 < 6 || LiIdx3 > 7){
        return(1);
    }

    //番号部６桁
    LsChar6 = LsWkContNo.substring(LiIdx2+1, 10);

    //ＰＲＥＦＩＸ部のデジット計算
    if(LsChar4 == "HLCU"){
        LlData = 84;          // 4 * 2^0 + 0 * 2^1 + 2 * 2^2 + 9 * 2^3
    }else{
        for (LiIdx1 = 1; LiIdx1 <= LiIdx2+1; LiIdx1++) {
            LsChar1 = LsWkContNo.substring(LiIdx1-1, LiIdx1);
 
            if (LsChar1 == "A") LiData1 = 10;
            if (LsChar1 == "B") LiData1 = 12;
            if (LsChar1 == "C") LiData1 = 13;
            if (LsChar1 == "D") LiData1 = 14;
            if (LsChar1 == "E") LiData1 = 15;
            if (LsChar1 == "F") LiData1 = 16;
            if (LsChar1 == "G") LiData1 = 17;
            if (LsChar1 == "H") LiData1 = 18;
            if (LsChar1 == "I") LiData1 = 19;
            if (LsChar1 == "J") LiData1 = 20;
            if (LsChar1 == "K") LiData1 = 21;
            if (LsChar1 == "L") LiData1 = 23;
            if (LsChar1 == "M") LiData1 = 24;
            if (LsChar1 == "N") LiData1 = 25;
            if (LsChar1 == "O") LiData1 = 26;
            if (LsChar1 == "P") LiData1 = 27;
            if (LsChar1 == "Q") LiData1 = 28;
            if (LsChar1 == "R") LiData1 = 29;
            if (LsChar1 == "S") LiData1 = 30;
            if (LsChar1 == "T") LiData1 = 31;
            if (LsChar1 == "U") LiData1 = 32;
            if (LsChar1 == "V") LiData1 = 34;
            if (LsChar1 == "W") LiData1 = 35;
            if (LsChar1 == "X") LiData1 = 36;
            if (LsChar1 == "Y") LiData1 = 37;
            if (LsChar1 == "Z") LiData1 = 38;
            snum = LsChar1.charCodeAt(1);
            if (snum < 65 || snum > 90){
                return(1);
            }
            LlData = LlData + LiData1 * Math.pow(2,(LiIdx1 - 1));
         
        }
    }
  
    //番号部分のデジット計算
    for (LiIdx1 = LiIdx2 + 1; LiIdx1 <= LiIdx2 + 6; LiIdx1++) {
        LsChar1 = LsWkContNo.substring(LiIdx1,LiIdx1+1);

        if (LsChar1 == "1") LiData1 = 1;
        if (LsChar1 == "2") LiData1 = 2;
        if (LsChar1 == "3") LiData1 = 3;
        if (LsChar1 == "4") LiData1 = 4;
        if (LsChar1 == "5") LiData1 = 5;
        if (LsChar1 == "6") LiData1 = 6;
        if (LsChar1 == "7") LiData1 = 7;
        if (LsChar1 == "8") LiData1 = 8;
        if (LsChar1 == "9") LiData1 = 9;
        if (LsChar1 == "0") LiData1 = 0;
        snum = LsChar1.charCodeAt(1);
        if (snum < 48 || snum > 57){
            return(1);
        }
      
        LlData = LlData + LiData1 * Math.pow(2,(LiIdx1));  
       
    }

              
    //チェックデジット値の算出
    LiAmari = LlData % 11;
    if(LiAmari == 10) LiAmari = 0;
   
    //チェックデジット付きコンテナ番号の生成
    LsChar1 = LsWkContNo.substring(LiIdx2+7, 11);    
    LsDigit = String(LiAmari);
    //入力コンテナ番号と計算したチェックデジットの比較
    if(LsChar1 != ""){
        if(LsChar1 == LsDigit){
            return(0);
        }else{
            return(9);
        }
    }else{
        return(1);
    }

}
//Y.TAKAKUWA Add-E 2015-03-13
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0  onload="finit();">
<!-------------実搬入番号入力画面--------------------------->
<% If ActionType = "M" Then %>


<TABLE border=0 cellPadding=3 cellSpacing=7 width="100%">
<TR>
<TD colspan=5 align=left>
<% If CheckDigit = "0" Then %>
<% '2016/11/17 H.Yoshikawa Del Start
'ブッキング番号、または、作業番号とコンテナ番号に<BR />			<!-- 2016/10/18 H.Yoshikawa Upd(作業番号の文言追加) -->
'入力間違いがないか、再度ご確認の上、次へ進んでください。
   '2016/11/17 H.Yoshikawa Del End %>
<% Else %>
<% '2016/11/17 H.Yoshikawa Upd Start
'ブッキング番号、または、作業番号とコンテナ番号に<BR />			<!-- 2016/10/18 H.Yoshikawa Upd(作業番号の文言追加) -->
'入力間違いがないか、再度ご確認の上、次へ進んでください。<BR/>
'<div style="color:red">※当コンテナ番号は一般海上輸送コンテナに該当しません。<BR/>
'コンテナ番号にお間違いが無ければ、次へ進んで下さい。</div>
%>
<div style="color:red">入力された当コンテナ番号は一般海上輸送コンテナに該当しませんでした。<BR/>
コンテナ番号にお間違いが無ければ『OK』ボタンを押して次へ進んでください。</div>
<% '2016/11/17 H.Yoshikawa Upd End %>
<% End If %>
</TD>
</TR>
<TR>
  <TD align=center>
    <input type="button" name="Send" value="   OK   " Onclick="fSend();" onkeypress="return true">
  </TD>
  <TD align=center>
    <input type="button" name="Stop" value="  修正  " Onclick="fStop();" onkeypress="return true">
  </TD>
</TR>
</TABLE>
<% Else %>
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD height="300" align=center>
<%'Mod-s 2006/03/06 h.matsuda%>
<!-----<FORM name="dmi310F" method="POST" action="./dmi315.asp">--->
      <FORM name="dmi310F" method="POST" action="./dmi312.asp">
	  <INPUT type=hidden name="ShoriMode" value="FLin">
<%' 2016/10/18 H.Yoshikawa Upd Start %>
<!--        <B>ブッキング番号</B><BR>
	  <INPUT type=text  name="BookNo" maxlength=20 size=27><BR>
        <B>コンテナ番号</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><P>
	  <A HREF="JavaScript:GoNext()">実行</A><P>
	  <A HREF="JavaScript:window.close()">閉じる</A><P>
-->
		<TABLE cellpadding=3>
		<TR>
			<TD colspan=2 style="border: 1px solid gray;">
		  		ブッキング番号、または、前回入力値を利用する<BR>
		  		作業番号を入力してください。<BR>
		  	</TD>
		</TR>
		<TR>
        	<TD><B>ブッキング番号</B></TD>
	  		<TD><INPUT type=text  name="BookNo" maxlength=20 size=27></TD>
	  	</TR>
		<TR>
        	<TD><B>または、作業番号</B></TD>
	  		<TD><INPUT type=text  name="WkNo" maxlength=5 size=10></TD>
	  	</TR>
		<TR>
			<TD colspan=2><BR></TD>
		</TR>
	  	<TR>
			<TD colspan=2 style="border: 1px solid gray;">
		  		コンテナ番号を入力してください。<BR>
		  	</TD>
		</TR>
        <TR>
        	<TD><B>コンテナ番号</B></TD>
	  		<TD><INPUT type=text  name="CONnum" maxlength=12></TD>
	  	</TR>
		<TR>
			<TD colspan=2><BR><BR></TD>
		</TR>
	  	<TR>
			<TD colspan=2 align="center">
				<A HREF="JavaScript:GoNext()">実行</A><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD colspan=2 align="center">
				<A HREF="JavaScript:window.close()">閉じる</A>
			</TD>
		</TR>
<%' 2016/10/18 H.Yoshikawa Upd End %>
      </FORM>
  </TD></TR>
</TABLE>
<%End If%>
<!-------------画面終わり--------------------------->
</BODY></HTML>
