<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi120.asp				_/
'_/	Function	:���O��������͉��			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/07/29	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���ύX	_/
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

'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'�f�[�^�擾
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
    WriteLogH "b202", "��������O������","12",""
  End If
'���O�C�����[�U�ɂ���ĉ�ЃR�[�h�X�V����
  saveCompCd CMPcd, UpFlag
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�����������</TITLE>
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
'�R���{�{�b�N�X�f�[�^�擾

'�R���e�i�T�C�Y�擾���\��
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

'�R���e�i�^�C�v�擾���\��
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

'�R���e�i�����擾���\��
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

'�R���e�i�ގ��擾���\��
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
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b202","03","������F�f�[�^����","102","�R���{�{�b�N�X�l�擾���s"
  end if

'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
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

//�R���e�i���
function GoConInfo(){
  Fname=document.dmi120F;
  ConInfo(Fname,1,0);
  return false;
}
//�o�^�E�X�V
function GoEntry(){
  target=document.dmi120F;
  <%'CW-034 If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then%>
  <% If Mord<>0 AND Request("TruckerFlag")<>1 AND UpFlag <> 1 Then%>
  if(target.way[1].checked){
    flag = confirm('�񓚂�No�ɂ��܂����H');
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
//�߂�
function GoBackT(){
  target=document.dmi120F;
  target.action="./dmi110.asp";
  return true;
}
//�폜
function GoDell(){
<%If Request("TruckerFlag")<>1 Then%>
  flag = confirm('�폜���܂����H');
<%Else%>
  flag = confirm('�w���悪����񓚍ςł��B\n�폜����O�Ɏw����Ɋm�F���Ă��������B\n�폜���܂����H');
<%End If%>
  if(flag){
    target=document.dmi120F;
    target.action="./dmi190.asp";
    return true;
  } else {
    return false;
  }
}
//���͏��`�F�b�N
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
        alert("���p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕�������͂��Ȃ��ł�������");
        strA[k].focus();
        return false;
      }
    }
  }
<% If UpFlag = 1 Then %>
  if(strA[0].value.length==0 && strA[4].value.length!=0){
    alert("�w��������ЂɎw�肵�Ȃ���΃w�b�hID����͂��鎖�͏o���܂���");
    strA[0].focus();
    return false;
  }
<% End If %>
  // Added 2003.8.3
  if(strA[4].value != ""){
    if(strA[4].value.length != 5){
      alert("�w�b�h�h�c�́u�w�b�h��ЃR�[�h�v�{�u�����R���v�œ��͂��Ă��������B");
      strA[4].focus();
      return false;
    }else{
      if(isNaN(strA[4].value.charAt(2)) || isNaN(strA[4].value.charAt(3)) || isNaN(strA[4].value.charAt(4))){
        alert("�w�b�h�h�c�́u�w�b�h��ЃR�[�h�v�{�u�����R���v�œ��͂��Ă��������B");
        strA[4].focus();
        return false;
      }
    }
  }
  // End of Addition 2003.8.3
  Num=LTrim(target.CONtear.value);
  if(Num.length==0){
    alert("�e�A�E�F�C�g���L�����Ă�������");
    target.CONtear.focus();
    return false;
  }
  ret = CheckSu(target.CONtear.value); 
  if(ret==false){
      alert("�����ȊO����͂��Ȃ��ł�������");
      target.CONtear.focus();
      return false;
  }
  ret = CheckSu(target.MaxW.value); 
  if(ret==false){
      alert("�����ȊO����͂��Ȃ��ł�������");
      target.MaxW.focus();
      return false;
  }
  strA    = new Array();
  strA[0] = target.CONsize;
  strA[1] = target.CONtype;
  strA[2] = target.CONhite;
  //strA[3] = target.CONsitu;				//-- 2016/10/24 H.Yoshikawa Del
  strM    = new Array("�T�C�Y","�^�C�v","����","�ގ�");
  for(k=0;k<strA.length;k++){
    if(strA[k].selectedIndex==0){
      alert(strM[k]+"��I�����Ă�������");
        strA[k].focus();
        return false;
    }
  }
<%' C-002 ADD START%>
  if(target.Comment1.value!="" && target.Comment1.value!=null){
    ret = CheckKin(target.Comment1.value); 
    if(ret==false){
      alert("�u\"�v��u\'�v���̔��p�L������͂��Ȃ��ł�������");
      target.Comment1.focus();
      return false;
    }
    retA=getByte(target.Comment1.value);
    if(retA[0]>70){
      if(retA[2]>35){
        alertStr="�S�p������5�����ȓ��œ��͂��Ă��������B";
      }else{
        alertStr="�S�p������"+Math.floor((70-retA[1])/2)+"�����ɂ��邩\n";
        alertStr=alertStr+"���p������"+(70-retA[2]*2)+"�����ɂ��Ă��������B";
      }
      alert("70�o�C�g�ȓ��œ��͂��Ă��������B\n70�o�C�g�ȓ��ɂ���ɂ�"+alertStr);
      target.Comment1.focus();
      return false;
    }
  }
<%' C-002 ADD END%>
<%' 3th ADD START%>
//���t�̃`�F�b�N
  if(!CheckDate('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Rmon,target.Rday,0))
      return false;
<%' 3th ADD End%>
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
<%'CW-017 ADD START%>
//�w�b�hID�̐���
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
  checkstr="���������������������������������������������������������������";
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
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi120F);finit();">
<!-------------����������͉��--------------------------->
<%=Request(CONnum)%>
<FORM name="dmi120F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
<% If Mord=0 Then %>
      <B>�����������</B>
<% Else %>
      <B>�����������(�X�V���[�h)</B>
<% End If %>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�R���e�i�m���D</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>" readOnly tabindex=-1></TD></TR>
  <TR>
    <TD width=230><BR><DIV class=bgb>��ЃR�[�h</DIV></TD>
    <TD>�o�^��<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" readOnly tabindex=-1 size=7>
        <INPUT type=text name="CMPcd1" value=<%=CMPcd(1)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value=<%=CMPcd(2)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value=<%=CMPcd(3)%> size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value=<%=CMPcd(4)%> size=5 maxlength=2>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�w�b�h�h�c</DIV></TD>
<!-- CW-017 Chenge
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5></TD></TR>
-->
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5 onBlur="checkIDF(1)"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�ԋp��</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" readOnly tabindex=-1></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����\���</DIV></TD>
<%'chage 3th    <TD><select name="Rmon" onchange="check_date('<%=DayTime(0)% >','<%=DayTime(1)% >',dmi021F.Rmon,dmi021F.Rday)">
'        </select>��<select name="Rday"></select>�� %>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=3 maxlength=2>��
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=3 maxlength=2>��
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>*�T�C�Y�A�^�C�v�A�����A�e�A�E�F�C�g</DIV></TD>
    <TD><select name="CONsize"></select>
        <select name="CONtype"></select>
        <select name="CONhite"></select>
        <select name="CONsitu" style="display:none;"></select>			<!-- 2016/10/24 H.Yoshikawa Upd (��\���Ƃ���) -->
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=5 maxlength=7>kg
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�戵�D��</DIV></TD>
    <TD><INPUT type=text name="TrhkSen" value="<%=Request("TrhkSen")%>" readOnly tabindex=-1 size=27></TD></TR>
  <TR>
    <TD><DIV class=bgb>�ۊ�</DIV></TD>
    <TD><select name="MrSk">
          <OPTION value=" "> 
          <OPTION value="Y">Y
          <OPTION value="N">N
        </select>
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>�l�`�w�d��</DIV></TD>
    <TD><INPUT type=text name="MaxW" value="<%=Request("MaxW")%>" maxlength=5>kg</TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>���l</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73 maxlength=70></TD></TR>
<%'C-002 ADD End %>

  <TR>
<!--  2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>�o�^�S����</DIV></TD>
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
       <INPUT type=submit value="�o�^" onClick="return GoEntry()">
       <INPUT type=submit value="�L�����Z��" onClick="window.close()">
<% Else %>

  <%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
    '  If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then
      If Request("TruckerFlag")<>1 AND UpFlag <> 1 AND Request("compFlag") Then%>
       <DIV class=bgw>�w�����։񓚁@�@�@
       <INPUT type=radio name="way" checked>Yes�@
       <INPUT type=radio name="way">No</DIV>
    </TD></TR>
    <TR><TD colspan=2 align=center>
  <% End If %>
  <%'20030909 IF Request("TruckerFlag")<>1 Then %>
  <% IF Request("TruckerFlag")<>1 AND Request("compFlag") Then %>
       <INPUT type=submit value="�X�V" onClick="return GoEntry()">
  <% End If %>
  <% IF UCase(Session.Contents("userid"))=CMPcd(0) Then %>
       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
       <INPUT type=submit value="�폜" onClick="return GoDell()">
  <% End If %>
       <INPUT type=submit value="�L�����Z��" onClick="window.close()">
<%'CW-023 Dell End If %>
<% End If 'CW-023 ADD%>
       <P>
       <INPUT type=submit value="�R���e�i���" onClick="return GoConInfo()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
