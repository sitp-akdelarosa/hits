<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi220.asp				_/
'_/	Function	:���O����o���͉��			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/08/06	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���S�ʉ��C	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
'�T�[�o���t�̎擾
 dim DayTime
 getDayTime DayTime
'�f�[�^����
  dim BookNo, COMPcd0, COMPcd1, Mord, TFlag
  dim Dflag,plintStr,i
  BookNo  = Request("BookNo")
  COMPcd0 = Request("COMPcd0")
  COMPcd1 = Request("COMPcd1")
  Mord    = Request("Mord")
  Dflag=""
  plintStr=""

  If Mord=0 Then '�V�K�o�^��
  
  Else          '�X�V��
    WriteLogH "b302", "����o���O������","12",""
    TFlag   = Request("TFlag")
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'    If COMPcd0 <> UCase(Session.Contents("userid")) OR TFlag = 1 Then
    If COMPcd0 <> UCase(Session.Contents("userid")) OR TFlag = 1 OR Request("compFlag")<>0 Then
      Dflag="readOnly"
    End If
    plintStr="(�X�V���[�h)"
  End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>����o�����͍X�V</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  window.resizeTo(550,680);
  bgset(target);
}
//�X�V
function GoNext(){
  target=document.dmi220F;
  if(!check(target))
    return;
  chengeUpper(target);
  target.action="./dmi230.asp";
  target.submit();
}
//�폜
function GoDell(){
<%If TFlag<>1 Then%>
  flag = confirm('�폜���܂����H');
<%Else%>
  flag = confirm('�w���悪����񓚍ςł��B\n�폜����O�Ɏw����Ɋm�F���Ă��������B\n�폜���܂����H');
<%End If%>
  if(flag){
    target=document.dmi220F;
    target.action="./dmi290.asp";
    len = target.elements.length;
    for (i=0; i<len; i++) target.elements[i].disabled = false;
    target.submit();
  }
}
//�ۗ�
function Suspend(){
  target=document.dmi220F;
  if(target.way[1].checked){
    flag = confirm('�񓚂�No�ɂ��܂����H');
    if(!flag) return false;
    target.Res.value=2;
  }
  target.action="./dmi230.asp";
  target.submit();
}
//�u�b�L���O���
function GoBookI(){
  target=document.dmi220F
  BookInfo(target);
}

//���͏��`�F�b�N
function check(target){
  if(!CheckEisu2(target.COMPcd1.value)){
    alert("��ЃR�[�h�ɔ��p�p�����ȊO�̕������L�����Ȃ��ł�������");
    target.COMPcd1.focus();
    return;
  }
  strA    = new Array();
  strA[0] = target.ContSize0;
  strA[1] = target.ContSize1;
  strA[2] = target.ContSize2;
  strA[3] = target.ContSize3;
  strA[4] = target.ContSize4;
  strA[5] = target.ContHeight0;
  strA[6] = target.ContHeight1;
  strA[7] = target.ContHeight2;
  strA[8] = target.ContHeight3;
  strA[9] = target.ContHeight4;
  strA[10]= target.PickNum0;
  strA[11]= target.PickNum1;
  strA[12]= target.PickNum2;
  strA[13]= target.PickNum3;
  strA[14]= target.PickNum4;
  strA[15]= target.vanMin;
  for(k=0;k<16;k++){
    if(strA[k].value!="" && strA[k].value!=null){
      ret = CheckSu(strA[k].value); 
      if(ret==false){
        alert("�����ȊO����͂��Ȃ��ł��������B");
        strA[k].focus();
        return false;
      }
    }
  }
  strA    = new Array();
  strA[0] = target.ContType0;
  strA[1] = target.ContType1;
  strA[2] = target.ContType2;
  strA[3] = target.ContType3;
  strA[4] = target.ContType4;
  strA[5] = target.Material0;
  strA[6] = target.Material1;
  strA[7] = target.Material2;
  strA[8] = target.Material3;
  strA[9] = target.Material4;
  for(k=0;k<10;k++){
    if(strA[k].value!="" && strA[k].value!=null){
      ret = CheckEisu2(strA[k].value); 
      if(ret==false){
        alert("���p�p�����ȊO�̕�������͂��Ȃ��ł�������");
        strA[k].focus();
        return false;
      }
    }
  }
//���t�̃`�F�b�N
  if(!CheckDate('<%=DayTime(0)%>','<%=DayTime(1)%>',target.vanMon,target.vanDay,target.vanHou)){
    return false;
  }else{
    if(target.vanHou.value=="")
      target.vanMin.value="";
    if(target.vanMin.value>59){
      alert("����0�`59�œ��͂��Ă�������");
      target.vanMin.focus();
      return false;
    }
  }
  NumA    = new Array();
  strA[0] = target.PickPlace0;	NumA[0]=20;
  strA[1] = target.PickPlace1;	NumA[1]=20;
  strA[2] = target.PickPlace2;	NumA[2]=20;
  strA[3] = target.PickPlace3;	NumA[3]=20;
  strA[4] = target.PickPlace4;	NumA[4]=20;
  strA[5] = target.vanPlace1;	NumA[5]=70;
  strA[6] = target.vanPlace2;	NumA[6]=70;
  strA[7] = target.goodsName;	NumA[7]=20;
  strA[8] = target.Comment1;	NumA[8]=70;
  strA[9] = target.Comment2;	NumA[9]=70;
  for(k=0;k<10;k++){
    if(strA[k].value!="" && strA[k].value!=null){
      ret = CheckKin(strA[k].value); 
      if(ret==false){
        alert("�u\"�v��u\'�v���̔��p�L������͂��Ȃ��ł��������B");
        strA[k].focus();
        return false;
      }
      retA=getByte(strA[k].value);
      if(retA[0]>NumA[k]){
        if(retA[2]>(NumA[k]/2)){
          alertStr="�S�p������"+(NumA[k]/2)+"�����ȓ��œ��͂��Ă��������B";
        }else{
          alertStr="�S�p������"+Math.floor((NumA[k]-retA[1])/2)+"�����ɂ��邩\n";
          alertStr=alertStr+"���p������"+(NumA[k]-retA[2]*2)+"�����ɂ��Ă��������B";
        }
        alert(NumA[k]+"�o�C�g�ȓ��œ��͂��Ă��������B\n"+NumA[k]+"�o�C�g�ȓ��ɂ���ɂ�"+alertStr);
        strA[k].focus();
        return false;
      }
    }
  }
  /* 2009/09/27 C.Pestano Del-S
   ret = CheckKana(target.TruckerSubName.value); 
   if(ret==false){
     alert("���p�J�i�����͓��͂ł��܂���");
     target.TruckerSubName.focus();
     return false;
   }2009/09/27 C.Pestano Del-E
   */
   
  return true;
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi220F.COMPcd1.focus();
}
//2008-01-31 Add-E M.Marquez

function CheckKana(str){
  checkstr="���������������������������������������������������������������";
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
		alert("���p��������͂��Ă��������B");
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
		// Shift_JIS: 0x0 �` 0x80, 0xa0  , 0xa1   �` 0xdf  , 0xfd   �` 0xff
		// Unicode  : 0x0 �` 0x80, 0xf8f0, 0xff61 �` 0xff9f, 0xf8f1 �` 0xf8f3
		if ( (c >= 0x0 && c < 0x81) || (c == 0xf8f0) || (c >= 0xff61 && c < 0xffa0) || (c >= 0xf8f1 && c < 0xf8f4)) {
			
		} else {			
			return false;		
		}
	}
	return true;
}
//2009/07/27 Add-E C.Pestano
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"  onLoad="setParam(document.dmi220F);finit();">
<!-------------����o�����͍X�V���--------------------------->
<FORM name="dmi220F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>����o������<%=plintStr%></B></TD></TR>
  <TR>
    <TD><DIV class=bgb>�u�b�L���O�m���D</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly tabindex=-1 size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly tabindex=-1 size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readOnly tabindex=-1 size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>�d���n</DIV></TD>
    <TD><INPUT type=text name="delivTo" value="<%=Request("delivTo")%>" readOnly tabindex=-1 size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>��ЃR�[�h(���^)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=Trim(COMPcd1)%>" size=5 <%=Dflag%> maxlength=2>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����Ɩ{��</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=0 cellSpacing=0 width=400 align=center>
      <TR><TD></TD><TD>�T�C�Y</TD><TD>�^�C�v</TD><TD>����</TD><TD>�ގ�</TD><TD>�s�b�N�ꏊ</TD><TD></TD><TD>�{��</TD></TR>
<% For i=0 To 4%>
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4 <%=Dflag%> maxlength=2></TD>
          <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4 <%=Dflag%> maxlength=2></TD>
          <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4 <%=Dflag%> maxlength=2></TD>
          <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>" size=4 <%=Dflag%> maxlength=1></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>" size=25 <%=Dflag%> maxlength=20></TD>
          <TD>�E�E�E</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4 <%=Dflag%> maxlength=3></TD></TR>
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߓ���</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3 <%=Dflag%> maxlength=2>��
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3 <%=Dflag%> maxlength=2>��
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3 <%=Dflag%> maxlength=2>��
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3 <%=Dflag%> maxlength=2>��
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�P</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�Q</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>�i��</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30 <%=Dflag%> maxlength=20></TD></TR>
  <TR>
    <TD><DIV class=bgb>������b�x�D�b�x�J�b�g��</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly tabindex=-1>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly tabindex=-1></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�P</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�Q</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>

  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>�o�^�S����</DIV></TD>
	<!-- 2009/07/25 Update C.Pestano -->
 	<TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerSubName")%>" maxlength=8 onBlur="CheckLen(this,true,true,false)"></TD></TR>
<!-- 2009/03/10 R.Shibuta Add-E -->
  <TR>
    <TD colspan=2 align=center>
<% If Request("ErrerM")<>"" Then %>
       <%= Request("ErrerM") %><BR>
<% Else %>
       <P><BR></P>
<% End If %>
       <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
<%'Add-s 2006/03/06 h.matsuda%>
       <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" >
	   <INPUT type=hidden name="ShoriMode" value="EMoutInf">
<%'Add-e 2006/03/06 h.matsuda%>
<% If Mord=0 Then %>
       <INPUT type=hidden name=Mord value="0" >
       <INPUT type=button value="�o�^" onClick="GoNext()">
<% ElseIf COMPcd0 = UCase(Session.Contents("userid")) Then%>
       <INPUT type=hidden name=Mord value="1" >
  <%If TFlag<>1 AND Request("compFlag")=0 Then%>
       <INPUT type=button value="�X�V" onClick="GoNext()">
  <% End If %>
       <INPUT type=button value="�폜" onClick="GoDell()">
<% Else %>
       <INPUT type=hidden name=Mord value="2" >
       <DIV class=bgw>�w�����։񓚁@�@�@
       <INPUT type=radio name="way" checked>Yes�@
       <INPUT type=radio name="way">No</DIV>
       <INPUT type=hidden name=Res value="1" >
    </TD></TR>
    <TR><TD colspan=2 align=center>
       <INPUT type=button value="�X�V" onClick="Suspend()">
<% End If %>
       <INPUT type=button value="�L�����Z��" onClick="window.close()">
       <P>
       <INPUT type=button value="�u�b�L���O���" onClick="GoBookI()">
    </TD></TR>


</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>