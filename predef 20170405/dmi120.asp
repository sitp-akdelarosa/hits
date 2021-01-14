<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi120.asp				_/
'_/	Function	:事前空搬入入力画面			_/
'_/	Date		:2003/05/28				_/
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

'サーバ日付の取得
 dim DayTime
 getDayTime DayTime

'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'データ取得
  dim UpFlag,Mord
  dim CONnum,CMPcd(5),Rmon,Rday,MrSk
  dim param,i,j
  Mord   = Request("Mord")
  CONnum = Request("CONnum")
  UpFlag = Request("UpFlag")
  For Each param In Request.Form
    If Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next
  Rmon = Request("Rmon")
  Rday = Request("Rday")
  MrSk = Request("MrSk")
  If Mord=2 Then Mord=1 End If
  If Mord=1 Then
    WriteLogH "b202", "空搬入事前情報入力","12",""
  End If
'ログインユーザによって会社コード更新制御
  saveCompCd CMPcd, UpFlag
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空搬入情報入力</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>

window.resizeTo(640,530);
<!--
function setParam(target){
<%'3th del
'3th del  setMonth(target.Rmon,'<%=Rmon% >');
'3th del  setDate(target.Rday,'<%=Rday% >');
'3th del  check_date('<%=DayTime(0)% >','<%=DayTime(1)% >',target.Rmon,target.Rday);%>
<%
'コンボボックスデータ取得

'コンテナサイズ取得＆表示
  StrSQL = "select * from mContSize ORDER BY ContSize ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContSize") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONsize,list,'" & Request("CONsize") & "');" & vbCrLf
  ObjRS.Close

'コンテナタイプ取得＆表示
  StrSQL = "select * from mContType ORDER BY ContType ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContType") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONtype,list,'" & Request("CONtype") & "');" & vbCrLf
  ObjRS.Close

'コンテナ高さ取得＆表示
  StrSQL = "select * from mContHeight ORDER BY ContHeight ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContHeight") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONhite,list,'" & Request("CONhite") & "');" & vbCrLf
  ObjRS.Close

'コンテナ材質取得＆表示
  StrSQL = "select * from mContMaterial ORDER BY ContMaterial ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContMaterial") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONsitu,list,'" & Request("CONsitu") & "');" & vbCrLf
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b202","03","空搬入：データ入力","102","コンボボックス値取得失敗"
  end if

'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0
%>
<%
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'  if(target.MrSk.options[1].value=="< %=MrSk% >"){
'    target.MrSk.selectedIndex=1;
'  } else if (target.MrSk.options[2].value=="< %=MrSk% >"){
'    target.MrSk.selectedIndex=2;
'  }
  If Mord=0 Then 
    Response.Write "  target.MrSk.selectedIndex=2;"&Chr(10)
  Else 
    Response.Write "  if(target.MrSk.options[1].value=="""&MrSk&"""){"&Chr(10)&_
                   "    target.MrSk.selectedIndex=1;"&Chr(10)&_
                   "  } else if (target.MrSk.options[2].value=="""&MrSk&"""){"&Chr(10)&_
                   "    target.MrSk.selectedIndex=2;"&Chr(10)&_
                   "  }"&Chr(10)
  End If
'Chang 20050303 End
%>
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
<%
'Change 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'  checkIDF(0);< %'CW-017 ADD% >
  If Mord=0 OR Request("compFlag") Then
    Response.Write "  checkIDF(0);"&Chr(10)
  End If
'Change 20050303 END
%>

}

//コンテナ情報
function GoConInfo(){
  Fname=document.dmi120F;
  ConInfo(Fname,1,0);
  return false;
}
//登録・更新
function GoEntry(){
  target=document.dmi120F;
  <%'CW-034 If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then%>
  <% If Mord<>0 AND Request("TruckerFlag")<>1 AND UpFlag <> 1 Then%>
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
  target.action="./dmi130.asp";
  chengeUpper(target);
  return true;
}
//戻る
function GoBackT(){
  target=document.dmi120F;
  target.action="./dmi110.asp";
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
    target=document.dmi120F;
    target.action="./dmi190.asp";
    return true;
  } else {
    return false;
  }
}
//入力情報チェック
function check(){
  target=document.dmi120F;
  strA    = new Array();
  strA[0] = target.CMPcd1;
  strA[1] = target.CMPcd2;
  strA[2] = target.CMPcd3;
  strA[3] = target.CMPcd4;
  strA[4] = target.HedId;
  for(k=0;k<strA.length;k++){
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
  Num=LTrim(target.CONtear.value);
  if(Num.length==0){
    alert("テアウェイトを記入してください");
    target.CONtear.focus();
    return false;
  }
  ret = CheckSu(target.CONtear.value); 
  if(ret==false){
      alert("数字以外を入力しないでください");
      target.CONtear.focus();
      return false;
  }
  ret = CheckSu(target.MaxW.value); 
  if(ret==false){
      alert("数字以外を入力しないでください");
      target.MaxW.focus();
      return false;
  }
  strA    = new Array();
  strA[0] = target.CONsize;
  strA[1] = target.CONtype;
  strA[2] = target.CONhite;
  //strA[3] = target.CONsitu;				//-- 2016/10/24 H.Yoshikawa Del
  strM    = new Array("サイズ","タイプ","高さ","材質");
  for(k=0;k<strA.length;k++){
    if(strA[k].selectedIndex==0){
      alert(strM[k]+"を選択してください");
        strA[k].focus();
        return false;
    }
  }
<%' C-002 ADD START%>
  if(target.Comment1.value!="" && target.Comment1.value!=null){
    ret = CheckKin(target.Comment1.value); 
    if(ret==false){
      alert("「\"」や「\'」等の半角記号を入力しないでください");
      target.Comment1.focus();
      return false;
    }
    retA=getByte(target.Comment1.value);
    if(retA[0]>70){
      if(retA[2]>35){
        alertStr="全角文字を5文字以内で入力してください。";
      }else{
        alertStr="全角文字を"+Math.floor((70-retA[1])/2)+"文字にするか\n";
        alertStr=alertStr+"半角文字を"+(70-retA[2]*2)+"文字にしてください。";
      }
      alert("70バイト以内で入力してください。\n70バイト以内にするには"+alertStr);
      target.Comment1.focus();
      return false;
    }
  }
<%' C-002 ADD END%>
<%' 3th ADD START%>
//日付のチェック
  if(!CheckDate('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Rmon,target.Rday,0))
      return false;
<%' 3th ADD End%>
	/* 2009/09/27 C.Pestano Del-S
   ret = CheckKana(target.TruckerSubName.value); 
   if(ret==false){
     alert("半角カナ文字は入力できません");
     target.TruckerSubName.focus();
     return false;
   }2009/09/27 C.Pestano Del-E
   */

  return true;
}
<%'CW-017 ADD START%>
//ヘッドIDの制御
function checkIDF(type){
<% 'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
   'If UpFlag <> 5 Then 
   If UpFlag <> 5 AND (Mord=0 OR Request("compFlag")) AND Request("TruckerFlag")<>1 Then%>
  target=document.dmi120F;
  targetCOMPcd=target.CMPcd<%=UpFlag%>;
  COMPcd="<%=Session.Contents("COMPcd")%>";
  checkID(type,target,targetCOMPcd,COMPcd);
<% End If %>
}
<%'CW-017 ADD END%>

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

//2008-01-30 Add-S M.Marquez
function finit(){
    document.dmi120F.CMPcd1.focus();
}
//2008-01-30 Add-E M.Marquez

// -->
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
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi120F);finit();">
<!-------------空搬入情報入力画面--------------------------->
<%=Request(CONnum)%>
<FORM name="dmi120F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
<% If Mord=0 Then %>
      <B>空搬入情報入力</B>
<% Else %>
      <B>空搬入情報入力(更新モード)</B>
<% End If %>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>コンテナＮｏ．</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>" readOnly tabindex=-1></TD></TR>
  <TR>
    <TD width=230><BR><DIV class=bgb>会社コード</DIV></TD>
    <TD>登録者<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" readOnly tabindex=-1 size=7>
        <INPUT type=text name="CMPcd1" value=<%=CMPcd(1)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value=<%=CMPcd(2)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value=<%=CMPcd(3)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value=<%=CMPcd(4)%> size=5 maxlength=2>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
<!-- CW-017 Chenge
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5></TD></TR>
-->
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5 onBlur="checkIDF(1)"></TD></TR>
  <TR>
    <TD><DIV class=bgb>返却先</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" readOnly tabindex=-1></TD></TR>
  <TR>
    <TD><DIV class=bgb>搬入予定日</DIV></TD>
<%'chage 3th    <TD><select name="Rmon" onchange="check_date('<%=DayTime(0)% >','<%=DayTime(1)% >',dmi021F.Rmon,dmi021F.Rday)">
'        </select>月<select name="Rday"></select>日 %>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=3 maxlength=2>月
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=3 maxlength=2>日
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>*サイズ、タイプ、高さ、テアウェイト</DIV></TD>
    <TD><select name="CONsize"></select>
        <select name="CONtype"></select>
        <select name="CONhite"></select>
        <select name="CONsitu" style="display:none;"></select>			<!-- 2016/10/24 H.Yoshikawa Upd (非表示とする) -->
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=5 maxlength=7>kg
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>取扱船社</DIV></TD>
    <TD><INPUT type=text name="TrhkSen" value="<%=Request("TrhkSen")%>" readOnly tabindex=-1 size=27></TD></TR>
  <TR>
    <TD><DIV class=bgb>丸関</DIV></TD>
    <TD><select name="MrSk">
          <OPTION value=" "> 
          <OPTION value="Y">Y
          <OPTION value="N">N
        </select>
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>ＭＡＸ重量</DIV></TD>
    <TD><INPUT type=text name="MaxW" value="<%=Request("MaxW")%>" maxlength=5>kg</TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>備考</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73 maxlength=70></TD></TR>
<%'C-002 ADD End %>

  <TR>
<!--  2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>登録担当者</DIV></TD>
	<!-- 2009/07/25 Update C.Pestano -->
 	<TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerName")%>" maxlength=8 onBlur="CheckLen(this,true,true,false)"></TD></TR>
<!--  2009/03/10 R.Shibuta Add-E -->
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name="UpUser"  value="<%=Request("UpUser")%>">
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
       <INPUT type=hidden name="compFlag"  value="<%=Request("compFlag")%>">
       <INPUT type=hidden name=Mord value="<%=Mord%>" >
<% If Mord=0 Then %>
       <INPUT type=submit value="登録" onClick="return GoEntry()">
       <INPUT type=submit value="キャンセル" onClick="window.close()">
<% Else %>

  <%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
    '  If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then
      If Request("TruckerFlag")<>1 AND UpFlag <> 1 AND Request("compFlag") Then%>
       <DIV class=bgw>指示元へ回答　　　
       <INPUT type=radio name="way" checked>Yes　
       <INPUT type=radio name="way">No</DIV>
    </TD></TR>
    <TR><TD colspan=2 align=center>
  <% End If %>
  <%'20030909 IF Request("TruckerFlag")<>1 Then %>
  <% IF Request("TruckerFlag")<>1 AND Request("compFlag") Then %>
       <INPUT type=submit value="更新" onClick="return GoEntry()">
  <% End If %>
  <% IF UCase(Session.Contents("userid"))=CMPcd(0) Then %>
       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
       <INPUT type=submit value="削除" onClick="return GoDell()">
  <% End If %>
       <INPUT type=submit value="キャンセル" onClick="window.close()">
<%'CW-023 Dell End If %>
<% End If 'CW-023 ADD%>
       <P>
       <INPUT type=submit value="コンテナ情報" onClick="return GoConInfo()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
