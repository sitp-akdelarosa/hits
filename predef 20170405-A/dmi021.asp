<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi021.asp									   _/
'_/	Function	:事前実搬出入力画面							   _/
'_/	Date		:2003/05/26									   _/
'_/	Code By		:SEIKO Electric.Co 大重						   _/
'_/	Modify		:C-002	2003/07/29	備考欄追加				   _/
'_/	Modify		:3th	2003/01/31	3次変更					   _/
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

'サーバ日付の取得
 dim DayTime
 getDayTime DayTime

'データ取得
  dim SakuNo,Flag,Num,CONnumA(),CMPcd(5),Rnissu
  dim param,i,j,Way,Mord,hCd,sUN,UpFlag,tmpstr
  sUN    = Session.Contents("sUN")
  SakuNo = Request("SakuNo")
  Flag   = Request("flag")
'3th del  Rmon = Request("Rmon")
'3th del  Rday = Request("Rday")
  Rnissu = Request("Rnissu")
  Num = Request("num")
  UpFlag = Request("UpFlag")
  ReDim CONnumA(Num)

  i=1
  For Each param In Request.Form
    tmpstr=""
    If Left(param, 6) = "CONnum" Then
      If param <> "CONnum" Then
        CONnumA(i) = Request.Form(param)
        tmpstr=tmpstr&CONnumA(i)
        i=i+1
      Else
        CONnumA(0) = Request.Form(param)
      End If
    ElseIf Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next
'ログインユーザによって会社コード更新制御
  saveCompCd CMPcd, UpFlag
 
'表示文言生成
  Way   =Array("","指定あり","指定なし","一覧から選択","ＢＬ番号")

  If SakuNo = "" Then '初期登録
    Mord = 0
    If Flag=3 Then
        WriteLogH "b105", "実搬出事前情報一覧(一覧から選択)","01",tmpstr
    End If
  Else                '更新
    Mord = 1
        WriteLogH "b10"&(2+Flag), "実搬出事前情報一覧("&Way(Flag)&")","12",""
  End If

'コンテナ番号受渡しメソッド
Sub Set_CONnum
  For i = 1 to Num-1
    Response.Write "       <INPUT type=hidden name='CONnum" & i & "' value='" & CONnumA(i) & "'>" & vbCrLf
  Next
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬出情報入力</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>

 window.resizeTo(600,730);

function setParam(target){
//  setMonth(target.Rmon,"<%'=Rmon%>");
//  setDate(target.Rday,"<%'=Rday%>");
  list = new Array("未入力","当日","2 日後","3 日後","4 日後","5 日後","5 日以上","リフトオフ")
  setList(target.Rnissu,list,"<%=Rnissu%>");
//  check_date('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Rmon,target.Rday)
  Utype=<%=Session.Contents("UType")%>;
  if(Utype != 5) target.HedId.readOnly = true;
<%
'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
  If Mord=1 AND (Request("TruckerFlag")=1 OR Not Request("compFlag")) Then
    Response.Write "  allsetreadOnly(target,8);"&Chr(10)
  End If
'ADD 20050303 END
%>
  bgset(target);
  checkIDF(0);<%'CW-017 ADD%>
}

//コンテナ情報
function GoConInfo(){
  Fname=document.dmi021F;
  ConInfo(Fname,<%=Flag%>,0);
  return false;
}
//登録・更新
function GoEntry(){
  target=document.dmi021F;
  <% If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then%>
  if(target.way[1].checked){
    flag = confirm('回答をNoにしますか？');
    if(!flag) return false;
    target.Mord.value=2;
  }
  <% End If %>
  ret = check();
  if(ret==false){
    return false;
  }
  target.action="./dmi030.asp";
  chengeUpper(target);
  return true;
}
//戻る
function GoBackT(){
  target=document.dmi021F;
  target.action="./dmi010.asp";
  return true;
}
//削除
function GoDell(){
<%If Request("TruckerFlag")<>1 Then%>
  flag = confirm('削除しますか？');
<%Else%>
  flag = confirm('指示先が受諾回答済です。\n削除する前に指示先に確認してください。\n削除しますか？');
<%End If%>
  if(flag){
    target=document.dmi021F;
    target.action="./dmi090.asp";
    return true;
  } else {
    return false;
  }
}
//入力情報チェック
function check(){
  target=document.dmi021F;
  strA    = new Array();
  strA[0] = target.CMPcd1;
  strA[1] = target.CMPcd2;
  strA[2] = target.CMPcd3;
  strA[3] = target.CMPcd4;
  strA[4] = target.HedId;
  strA[5] = target.HTo;

  
  for(k=0;k<6;k++){
    if(strA[k].value!="" && strA[k].value!=null && strA[k].readOnly==false){
      ret = CheckEisu(strA[k].value); 
      if(ret==false){
        alert("半角英数字と半角スペース、「-」、「/」以外の文字を入力しないでください");
        strA[k].focus();
        return false;
      }
    }
  }

<% If UpFlag = 1 Then %>
  if(strA[0].value.length==0 && strA[4].value.length!=0){
    alert("指示先を自社に指定しなければヘッドIDを入力する事は出来ません");
    strA[0].focus();
    return false;
  }
<% End If %>
  // Added 2003.8.3
  if(strA[4].value != ""){
    if(strA[4].value.length != 5){
      alert("ヘッドＩＤは「ヘッド会社コード」＋「数字３桁」で入力してください。");
      strA[4].focus();
      return false;
    }else{
      if(isNaN(strA[4].value.charAt(2)) || isNaN(strA[4].value.charAt(3)) || isNaN(strA[4].value.charAt(4))){
        alert("ヘッドＩＤは「ヘッド会社コード」＋「数字３桁」で入力してください。");
        strA[4].focus();
        return false;
      }
    }
  }
  // End of Addition 2003.8.3
<%' C-002 ADD START%>
  NumA    = new Array();
  strA[0] = target.Nonyu1;		NumA[0]=70;
  strA[1] = target.Nonyu2;		NumA[1]=70;
  strA[2] = target.Comment1;	NumA[2]=70;
  strA[3] = target.Comment2;	NumA[3]=70;
  strA[4] = target.HinName;		NumA[4]=20;
  for(k=0;k<5;k++){
    if(strA[k].value!="" && strA[k].value!=null){
      ret = CheckKin(strA[k].value); 
      if(ret==false){
        alert("「\"」や「\'」等の半角記号を入力しないでください。");
        strA[k].focus();
        return false;
      }
      retA=getByte(strA[k].value);
      if(retA[0]>NumA[k]){
        if(retA[2]>(NumA[k]/2)){
          alertStr="全角文字を"+(NumA[k]/2)+"文字以内で入力してください。";
        }else{
          alertStr="全角文字を"+Math.floor((NumA[k]-retA[1])/2)+"文字にするか\n";
          alertStr=alertStr+"半角文字を"+(NumA[k]-retA[2]*2)+"文字にしてください。";
        }
        alert(NumA[k]+"バイト以内で入力してください。\n"+NumA[k]+"バイト以内にするには"+alertStr);
        strA[k].focus();
        return false;
      }
    }
  }
   /* 2009/09/27 C.Pestano Del-S
   ret = CheckKana(target.TruckerSubName.value); 
   if(ret==false){
     alert("半角カナ文字は入力できません");
     target.TruckerSubName.focus();
     return false;
   }  2009/09/27 C.Pestano Del-E
   */
   
<%' C-002 ADD END%>
<%' 3th ADD START%>
//日付のチェック
  if(!CheckDate('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Rmon,target.Rday,target.Rhou))
      return false;
  <!-- 2008/01/31 Edit S G.Ariola -->
  if(!CheckDatewithMin('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Nomon,target.Noday,target.Nohou,target.Nomin))
  <!-- 2008/01/31 Edit E G.Ariola -->
    return false;
<%' 3th ADD End%>
  return true;
}
<%'CW-017 ADD START%>
//ヘッドIDの制御
function checkIDF(type){
<% 'Change 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
   'If UpFlag <> 5 Then 
   If UpFlag <> 5 AND (Mord=0 OR Request("compFlag")) AND Request("TruckerFlag")<>1 Then%>
  target=document.dmi021F;
  targetCOMPcd=target.CMPcd<%=UpFlag%>;
  COMPcd="<%=Session.Contents("COMPcd")%>";
  checkID(type,target,targetCOMPcd,COMPcd);
<% End If %>
}
<%'CW-017 ADD END%>

//2008-01-31 Add-S G.Ariola
function finit(){
    document.dmi021F.HinName.focus();
}
//2008-01-31 Add-E G.Ariola
// -->

function CheckKana(str){
  checkstr="｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
   for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) >= 0){
      return false;
    }
  }
  return true;
}
//2009/07/27 Add-S C.Pestano
function CheckLen(obj,mesgon,focuson,mandatory) {
	var kanjicheck = gfStrLen(obj.value);
	
	if (kanjicheck == false){
		alert("半角文字を入力してください。");
		obj.focus();
		return false;
	}	
	
	if (mandatory && objlength==0)
		return false;	
	return true;
}

function gfStrLen(StrSrc) {
	var r = 0;
	for (var i = 0; i < StrSrc.length; i++) {
		var c = StrSrc.charCodeAt(i);
		// Shift_JIS: 0x0 ～ 0x80, 0xa0  , 0xa1   ～ 0xdf  , 0xfd   ～ 0xff
		// Unicode  : 0x0 ～ 0x80, 0xf8f0, 0xff61 ～ 0xff9f, 0xf8f1 ～ 0xf8f3
		if ( (c >= 0x0 && c < 0x81) || (c == 0xf8f0) || (c >= 0xff61 && c < 0xffa0) || (c >= 0xf8f1 && c < 0xf8f4)) {
			
		} else {			
			return false;		
		}
	}
	return true;
}
//2009/07/27 Add-E C.Pestano
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi021F);finit();">
<!-------------実搬出情報入力画面--------------------------->
<FORM name="dmi021F" method="POST" scrolling="yes">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
<% If Mord=0 Then %>
    <TD colspan=2>
      <B>実搬出情報入力</B></TD></TR>
<% Else %>
    <TD><B>実搬出情報入力(更新モード)</B></TD>
    <TD><TABLE border=1 cellPadding=3 cellSpacing=0 align="right">
          <TR bgcolor="#f0f0f0"><TD>作業番号</TD><TD><%=SakuNo%></TD></TR>
        </TABLE>
        <INPUT type=hidden name="SakuNo"  value="<%=SakuNo%>">
    </TD></TR>
<% End If %>
  <TR>
<% If Flag=4 Then %>
    <TD><DIV class=bgb>ＢＬＮｏ．</DIV></TD>
    <TD><INPUT type=text name="BLnum" value="<%=Request("BLnum")%>" readOnly tabindex=-1>　　<%=Way(Flag)%>
        <INPUT type=hidden name="CONnum" value="<%=CONnumA(0)%>"></TD></TR>
<% Else %>
    <TD><DIV class=bgb>コンテナＮｏ．</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnumA(0)%>" readOnly tabindex=-1>　　<%=Way(Flag)%>
        <INPUT type=hidden name="BLnum" value="<%=Request("BLnum")%>"></TD></TR>
<% End If %>
  <TR>
    <TD width=180>
        <DIV class=bgb>サイズ、タイプ、高さ、グロス</DIV></TD>
    <TD><INPUT type=text name="CONsize" value="<%=Request("CONsize")%>" readOnly tabindex=-1 size=5>
        <INPUT type=text name="CONtype" value="<%=Request("CONtype")%>" readOnly tabindex=-1 size=5>
        <INPUT type=text name="CONhite" value="<%=Request("CONhite")%>" readOnly tabindex=-1 size=5>
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" readOnly tabindex=-1 size=5>kg
    </TD></TR>
<%'3th追加 Start%>
  <TR>
    <TD><DIV class=bgb>船社、船名</DIV></TD>
    <TD><INPUT type=text name="Shipfact" value="<%=Request("shipFact")%>" readOnly tabindex=-1 size=20>
        <INPUT type=text name="ShipName" value="<%=Request("shipName")%>" readOnly tabindex=-1 size=20>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>品名</DIV></TD>
    <TD><INPUT type=text name="HinName" value="<%=Request("HinName")%>" tabindex=0 size=40 maxlength=20>
    </TD></TR>
<%'3th追加 End%>
  <TR>
    <TD><BR><DIV class=bgb>会社コード</DIV></TD>
    <TD>登録者<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" readOnly tabindex=-1 size=7>
        <INPUT type=text name="CMPcd1" value=<%=CMPcd(1)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value=<%=CMPcd(2)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value=<%=CMPcd(3)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value=<%=CMPcd(4)%> size=5 maxlength=2>
<%'CW-017 ADD ENDT%>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
<!-- CW-017 Chenge
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5></TD></TR>
-->
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5 onBlur="checkIDF(1)"></TD></TR>
  <TR>
    <TD><DIV class=bgb>ＣＹ</DIV></TD>
    <TD><INPUT type=text name="HFrom" value="<%=Request("HFrom")%>" readOnly tabindex=-1 ></TD></TR>
  <TR>
    <TD><DIV class=bgb>搬出予定日</DIV></TD>
<%'chage 3th    <TD><select name="Rmon" onchange="check_date('<%=DayTime(0)% >','<%=DayTime(1)% >',dmi021F.Rmon,dmi021F.Rday)">
'        </select>月<select name="Rday"></select>日 %>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=3 maxlength=2>月
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=3 maxlength=2>日
        <INPUT type=text name="Rhou" value="<%=Request("Rhou")%>" size=3 maxlength=2>時
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>搬出先</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" size=30 maxlength=26></TD></TR>
<%'3th追加 Start%>
  <TR>
    <TD><DIV class=bgb>納入先１</DIV></TD>
    <TD><INPUT type=text name="Nonyu1" value="<%=Request("Nonyu1")%>" size=70 maxlength=70>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>納入先２</DIV></TD>
    <TD><INPUT type=text name="Nonyu2" value="<%=Request("Nonyu2")%>" size=70 maxlength=70>
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
    <TD><INPUT type=text name="RPlace" value="<%=Request("RPlace")%>" size=30  readOnly tabindex=-1>
    </TD></TR>
<%'3th追加 End%>
  <TR>
    <TD><DIV class=bgb>返却予定日数（フリータイム）</DIV></TD>
    <TD><select name="Rnissu"></select>
    </TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>備考１</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73 maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考２</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73 maxlength=70></TD></TR>
    
  <TR>
  
<!-- 2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>登録担当者</DIV></TD>
	<!-- 2009/07/27 Update C.Pestano -->
 	<TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerName")%>" maxlength=8 onBlur="CheckLen(this,true,true,false)"></TD></TR>
<!-- 2009/03/10 R.Shibuta Add-E -->

<%'Del 3th  <TR>
'    <TD><DIV class=bgb>備考３</DIV></TD>
'    <TD><INPUT type=text name="Comment3" value="<%=Request("Comment3")% >" size=13 maxlength=10></TD></TR>%>
<%'C-002 ADD Start %>
<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
  ' If Mord<>0 AND Request("TruckerFlag")<>1 AND UpFlag <> 1 Then


  
 If Mord<>0 AND Request("TruckerFlag")<>1 AND UpFlag <> 1 AND Request("compFlag") Then %>
  <TR>
    <TD colspan=2 align=center>
       <DIV class=bgw>指示元へ回答　　　
       <INPUT type=radio name="way" checked>Yes　
       <INPUT type=radio name="way">No</DIV>
    </TD></TR>
<% End If %>


  
  <TR>
    <TD colspan=2 align=center>
       <DIV class=alert><B>＜注意＞</B>船社によってはゲートで認証IDの入力が必要になります</DIV>
    </TD></TR>
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name="UpUser"  value="<%=Request("UpUser")%>">
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
       <INPUT type=hidden name="compFlag"  value="<%=Request("compFlag")%>">
       <INPUT type=hidden name=flag value="<%=Flag%>" >
       <INPUT type=hidden name=num value="<%=Num%>" >
       <INPUT type=hidden name=Mord value="1" >
<% If Num > 1 Then call Set_CONnum End If%>
<% If Mord=0 Then %>
       <INPUT type=submit value="登録" onClick="return GoEntry()">
       <INPUT type=submit value="キャンセル" onClick="window.close()">
       <INPUT type=submit value="コンテナ情報" onClick="return GoConInfo()">
<% Else %>
  <%'20030909 IF Request("TruckerFlag")<>1 Then %>
  <% IF Request("TruckerFlag")<>1 AND Request("compFlag") Then %>
       <INPUT type=submit value="更新" onClick="return GoEntry()">
  <% End If %>
  <% IF UCase(Session.Contents("userid"))=CMPcd(0) Then %>
       <INPUT type=submit value="削除" onClick="return GoDell()">
       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
  <% End If %>
       <INPUT type=submit value="キャンセル" onClick="window.close()">
<% End If %>
       <P>
       <INPUT type=submit value="コンテナ情報" onClick="return GoConInfo()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
<SCRIPT language=JavaScript>
setParam(document.dmi021F);
</SCRIPT>
</BODY></HTML>
