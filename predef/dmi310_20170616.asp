<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi310.asp				_/
'_/	Function	:���O�������ԍ����͉��		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:3th	2003/01/31	3���ύX	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%><% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b402", "���������O������","00",""
  
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
<TITLE><%If ActionType <> "M" Then %>���O�o�^�E�����[�쐬<%End If%></TITLE>
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
  strA    = new Array("�u�b�L���O�ԍ�","�R���e�i�ԍ�","��Ɣԍ�");		// 2016/10/18 H.Yoshikawa Upd�i��Ɣԍ��ǉ��j
  target=document.dmi310F;

// 2016/10/18 H.Yoshikawa Add Start
  if(Rtrim(target.BookNo.value, " ")=="" && Rtrim(target.WkNo.value, " ")==""
   || Rtrim(target.BookNo.value, " ")!="" && Rtrim(target.WkNo.value, " ")!=""){
      alert("�u�b�L���O�ԍ��A�܂��́A��Ɣԍ������ꂩ���L�����Ă�������");
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
      alert(strA[k]+"���L�����Ă�������");
      targetA[k].focus();
      return;
    }
   }																	// 2016/10/18 H.Yoshikawa Add
    if(k==0){
      if(!CheckEisu(targetA[k].value)){
        alert(strA[k]+"�ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕������L�����Ȃ��ł�������");
        targetA[k].focus();
        return;
      }
    }else{
      if(!CheckEisu2(targetA[k].value)){
        alert(strA[k]+"�ɔ��p�p�����ȊO�̕������L�����Ȃ��ł�������");
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
//  �@�\   : �R���e�i�ԍ��̃f�B�W�b�g�`�F�b�N���s��
//
//  ����   : sContNo           As String     - [I] �R���e�i�ԍ�
//
//  �߂�l �F�`�F�b�N����
//             0 - ����
//             1 - �v�Z�s�\�R���e�i
//             9 - �`�F�b�N�f�B�W�b�g�G���[
//            -1 - ��O�G���[
//**************************************************
function gfJDigitCheck(sContNo){

    var LsChar1;    //�P�����G���A
    var LsChar4;    //�S�����G���A
    var LsChar6;    //�U�����G���A
    var LsWkContNo; //�R���e�i�m�n�i�啶���j
    var LiIdx1;     //�Y��
    var LiIdx2;     //�Y��
    var LiIdx3;     //�Y��
    var LiData1;    //�v�Z�G���A
    var LiAmari;    //�v�Z�G���A
    var LiLen;      //����
    var LlData = 0;     //�v�Z�G���A
    var LsDigit;    
    var snum;    
    
    LiIdx2 = 0;
    LiIdx3 = 0;
    LsWkContNo = sContNo.value.toUpperCase();
       
    LiLen = sContNo.value.length;
    
    //���͂���Ȃ��`�F�b�N
    if(LiLen==0){
        return(1);
    }
     
    for (LiIdx1 = 1; LiIdx1 <= LiLen; LiIdx1++) {
        //���e�������̕ϊ��R�[�h�g�p
        //65: "A" �` 90: "Z"  
        snum = LsWkContNo.charCodeAt(LiIdx1);
        if (snum >= 65 && snum <= 90){
            LiIdx2 = LiIdx2 + 1;
        }else{ 
            break;
        }
    }
    
    //�o�q�d�e�h�w�̑Ó����`�F�b�N
    if(LiIdx2 == 0 || LiIdx2 < 3 ){
        return(1);
    }
 
    LsChar4 = LsWkContNo.substring(0, LiIdx2);

    //�ԍ����U���`�F�b�N 
    //48: "0" �` 57: "9"  
    for (LiIdx1 = LiIdx2 + 1; LiIdx1 <= 12; LiIdx1++) {
        //���e�������̕ϊ��R�[�h�g�p
        //48: "0" �` 57: "9"  
        snum = LsWkContNo.charCodeAt(LiIdx1);
        if (snum >= 48 && snum <= 57){
            LiIdx3 = LiIdx3 + 1;
        }else{ 
            break;
        }
    }
    //�ԍ����U�`�V���ȊO�G���[
    if(LiIdx3 < 6 || LiIdx3 > 7){
        return(1);
    }

    //�ԍ����U��
    LsChar6 = LsWkContNo.substring(LiIdx2+1, 10);

    //�o�q�d�e�h�w���̃f�W�b�g�v�Z
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
  
    //�ԍ������̃f�W�b�g�v�Z
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

              
    //�`�F�b�N�f�W�b�g�l�̎Z�o
    LiAmari = LlData % 11;
    if(LiAmari == 10) LiAmari = 0;
   
    //�`�F�b�N�f�W�b�g�t���R���e�i�ԍ��̐���
    LsChar1 = LsWkContNo.substring(LiIdx2+7, 11);    
    LsDigit = String(LiAmari);
    //���̓R���e�i�ԍ��ƌv�Z�����`�F�b�N�f�W�b�g�̔�r
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
<!-------------�������ԍ����͉��--------------------------->
<% If ActionType = "M" Then %>


<TABLE border=0 cellPadding=3 cellSpacing=7 width="100%">
<TR>
<TD colspan=5 align=left>
<% If CheckDigit = "0" Then %>
<% '2016/11/17 H.Yoshikawa Del Start
'�u�b�L���O�ԍ��A�܂��́A��Ɣԍ��ƃR���e�i�ԍ���<BR />			<!-- 2016/10/18 H.Yoshikawa Upd(��Ɣԍ��̕����ǉ�) -->
'���͊ԈႢ���Ȃ����A�ēx���m�F�̏�A���֐i��ł��������B
   '2016/11/17 H.Yoshikawa Del End %>
<% Else %>
<% '2016/11/17 H.Yoshikawa Upd Start
'�u�b�L���O�ԍ��A�܂��́A��Ɣԍ��ƃR���e�i�ԍ���<BR />			<!-- 2016/10/18 H.Yoshikawa Upd(��Ɣԍ��̕����ǉ�) -->
'���͊ԈႢ���Ȃ����A�ēx���m�F�̏�A���֐i��ł��������B<BR/>
'<div style="color:red">�����R���e�i�ԍ��͈�ʊC��A���R���e�i�ɊY�����܂���B<BR/>
'�R���e�i�ԍ��ɂ��ԈႢ��������΁A���֐i��ŉ������B</div>
%>
<div style="color:red">���͂��ꂽ���R���e�i�ԍ��͈�ʊC��A���R���e�i�ɊY�����܂���ł����B<BR/>
�R���e�i�ԍ��ɂ��ԈႢ��������΁wOK�x�{�^���������Ď��֐i��ł��������B</div>
<% '2016/11/17 H.Yoshikawa Upd End %>
<% End If %>
</TD>
</TR>
<TR>
  <TD align=center>
    <input type="button" name="Send" value="   OK   " Onclick="fSend();" onkeypress="return true">
  </TD>
  <TD align=center>
    <input type="button" name="Stop" value="  �C��  " Onclick="fStop();" onkeypress="return true">
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
<!--        <B>�u�b�L���O�ԍ�</B><BR>
	  <INPUT type=text  name="BookNo" maxlength=20 size=27><BR>
        <B>�R���e�i�ԍ�</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><P>
	  <A HREF="JavaScript:GoNext()">���s</A><P>
	  <A HREF="JavaScript:window.close()">����</A><P>
-->
		<TABLE cellpadding=3>
		<TR>
			<TD colspan=2 style="border: 1px solid gray;">
		  		�u�b�L���O�ԍ��A�܂��́A�O����͒l�𗘗p����<BR>
		  		��Ɣԍ�����͂��Ă��������B<BR>
		  	</TD>
		</TR>
		<TR>
        	<TD><B>�u�b�L���O�ԍ�</B></TD>
	  		<TD><INPUT type=text  name="BookNo" maxlength=20 size=27></TD>
	  	</TR>
		<TR>
        	<TD><B>�܂��́A��Ɣԍ�</B></TD>
	  		<TD><INPUT type=text  name="WkNo" maxlength=5 size=10></TD>
	  	</TR>
		<TR>
			<TD colspan=2><BR></TD>
		</TR>
	  	<TR>
			<TD colspan=2 style="border: 1px solid gray;">
		  		�R���e�i�ԍ�����͂��Ă��������B<BR>
		  	</TD>
		</TR>
        <TR>
        	<TD><B>�R���e�i�ԍ�</B></TD>
	  		<TD><INPUT type=text  name="CONnum" maxlength=12></TD>
	  	</TR>
		<TR>
			<TD colspan=2><BR><BR></TD>
		</TR>
	  	<TR>
			<TD colspan=2 align="center">
				<A HREF="JavaScript:GoNext()">���s</A><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD colspan=2 align="center">
				<A HREF="JavaScript:window.close()">����</A>
			</TD>
		</TR>
<%' 2016/10/18 H.Yoshikawa Upd End %>
      </FORM>
  </TD></TR>
</TABLE>
<%End If%>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
