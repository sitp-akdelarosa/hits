<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi030.asp				_/
'_/	Function	:���O�����o���͊m�F���			_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/07/29	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���ύX	_/
'_/	Modify		:		2018/10/16	��ƃ��[������`�F�b�N�ǉ�	_/
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

'�f�[�^�擾
  dim SakuNo,Flag,Num,CONnumA(),CMPcd(5),Rmon,Rday
  dim param,i,j,Way,Mord,tmpstr,tmpNo
  SakuNo = Request("SakuNo")
  Flag= Request("flag")
  Num = Request("num")
  ReDim CONnumA(Num)
  i=0
  For Each param In Request.Form
    If Left(param, 6) = "CONnum" Then
      If param <> "CONnum" Then
        i = Mid(param,7)
        CONnumA(i) = Request.Form(param)
      Else
        CONnumA(0) = Request.Form(param)
      End If
    ElseIf Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next

'�\����������
'3th del  If Request("Rmon") = 0 Then 
'3th del    Rmon = " "
'3th del  Else
'3th del    Rmon = Right("0"&Request("Rmon"),2)
'3th del  End If
'3th del  If Request("Rday") = 0 Then 
'3th del    Rday = " "
'3th del  Else
'3th del    Rday = Right("0"&Request("Rday"),2)
'3th del  End If
  Way   =Array("","�w�肠��","�w��Ȃ�","�ꗗ����I��","�a�k�ԍ�")
  If SakuNo = "" Then '�����o�^
    Mord = 0
    tmpNo="02"
  Else                '�X�V
    Mord = Request("Mord")
    tmpNo="13"
  End If

  dim ret
  If Mord=2 Then
    ret = true
  Else
  '�G���[�g���b�v�J�n
    on error resume next
  'DB�ڑ�
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS
  '�w�b�hID�̃`�F�b�N
    checkHdCd ObjConn, ObjRS, CMPcd, ret
  'DB�ڑ�����
    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0
  End If
  If Request("UpFlag") <> 5 Then 
    tmpstr=CMpcd(Request("UpFlag"))&"/"
  Else
    tmpstr="/"
  End If
  tmpstr=tmpstr&Request("HedId")&"/"&Request("HTo")&"/"&Rmon&Rday&_
         "/"&Request("Rnissu")
  If ret Then
    tmpstr=tmpstr&",���͓��e�̐���:0(������)"
  Else
    tmpstr=tmpstr&",���͓��e�̐���:1(���)"
  End If
  WriteLogH "b10"&(2+Flag), "�����o���O���ꗗ("&Way(Flag)&")", tmpNo,tmpstr

'�R���e�i�ԍ���n�����\�b�h
Sub Set_CONnum
  For i = 1 to Num -1
    Response.Write "       <INPUT type=hidden name='CONnum" & i & "' value='" & CONnumA(i) & "'>" & vbCrLf
  Next
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�����o�����͊m�F</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//��ʕ\��
function setParam(target){
  len = target.elements.length;
  for (i=0; i<len-3; i++) target.elements[i].readOnly = true;
  bgset(target);
}

//�o�^
function GoEntry(){
  target=document.dmi030F;
  //2018/10/16 H.Yoshikawa Add-S
  if(target.ChkRule){
    if(!target.ChkRule.checked){
      alert('�u���[�h�����ӎ��������炵�A���S��Ƃ��s���܂��v�Ƀ`�F�b�N�����i���񂵂āj�\����������ĉ������B���������Ȃ��ꍇ�A��Ɨ\��o���܂���B');
      return false;
    }
  }
  //2018/10/16 H.Yoshikawa Add-E
  target.action="./dmi040.asp";
  return true;
}
//�߂�
function GoBackT(){
  target=document.dmi030F;
  target.action="./dmi021.asp";
  return true;
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�����o�����͊m�F���--------------------------->
<FORM name="dmi030F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
<% If Mord = 0 Then %>
    <TD colspan=2>
      <B>�����o�����͊m�F</B></TD></TR>
<% Else %>
    <TD><B>�����o�����͊m�F</B></TD>
    <TD><TABLE border=1 cellPadding=3 cellSpacing=0 align="right">
          <TR bgcolor="#f0f0f0"><TD>��Ɣԍ�</TD><TD><%=SakuNo%></TD></TR>
        </TABLE>
        <INPUT type=hidden name="SakuNo"  value="<%=SakuNo%>">
    </TD></TR>
<% End If %>
  <TR>
<% If Flag=4 Then %>
    <TD><DIV class=bgb>�a�k�m���D</DIV></TD>
    <TD><INPUT type=text name="BLnum" value="<%=Request("BLnum")%>">�@�@<%=Way(Flag)%>
    <INPUT type=hidden name="CONnum" value="<%=CONnumA(0)%>"></TD></TR>
<% Else %>
    <TD><DIV class=bgb>�R���e�i�m���D</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnumA(0)%>">�@�@<%=Way(Flag)%></TD></TR>
        <INPUT type=hidden name="BLnum"   value="<%=Request("BLnum")%>">
<% End If %>
  <TR>
    <TD width=180>
        <DIV class=bgb>�T�C�Y�A�^�C�v�A�����A�O���X</DIV></TD>
    <TD><INPUT type=text name="CONsize" value="<%=Request("CONsize")%>" size=5>
        <INPUT type=text name="CONtype" value="<%=Request("CONtype")%>" size=5>
        <INPUT type=text name="CONhite" value="<%=Request("CONhite")%>" size=5>
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=5>kg
    </TD></TR>
<%'3th�ǉ� Start%>
  <TR>
    <TD><DIV class=bgb>�D�ЁA�D��</DIV></TD>
    <TD><INPUT type=text name="Shipfact" value="<%=Request("shipFact")%>" size=20>
        <INPUT type=text name="ShipName" value="<%=Request("shipName")%>" size=20>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�i��</DIV></TD>
    <TD><INPUT type=text name="HinName" value="<%=Request("HinName")%>" size=40 maxlength=20>
    </TD></TR>
<%'3th�ǉ� End%>
  <TR>
    <TD><BR><DIV class=bgb>��ЃR�[�h</DIV></TD>
    <TD>�o�^��<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5 maxlength=2>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�w�b�h�h�c</DIV></TD>
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�b�x</DIV></TD>
    <TD><INPUT type=text name="HFrom" value="<%=Request("Hfrom")%>"></TD></TR>
    <TD><DIV class=bgb>���o�\���</DIV></TD>
<%'chage 3th    <TD><select name="Rmon" onchange="check_date('<%=DayTime(0)% >','<%=DayTime(1)% >',dmi021F.Rmon,dmi021F.Rday)">
'        </select>��<select name="Rday"></select>�� %>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=3 maxlength=2>��
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=3 maxlength=2>��
        <INPUT type=text name="Rhou" value="<%=Request("Rhou")%>" size=3 maxlength=2>��
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>���o��</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" size=30></TD></TR>
<%'3th�ǉ� Start%>
  <TR>
    <TD><DIV class=bgb>�[����P</DIV></TD>
    <TD><INPUT type=text name="Nonyu1" value="<%=Request("Nonyu1")%>" size=73>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�[����Q</DIV></TD>
    <TD><INPUT type=text name="Nonyu2" value="<%=Request("Nonyu2")%>" size=73>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�[������</DIV></TD>
    <TD><INPUT type=text name="Nomon" value="<%=Request("Nomon")%>" size=3 maxlength=2>��
        <INPUT type=text name="Noday" value="<%=Request("Noday")%>" size=3 maxlength=2>��
        <INPUT type=text name="Nohou" value="<%=Request("Nohou")%>" size=3 maxlength=2>��
		<!-- 2008/01/31 Add S G.Ariola -->
		<INPUT type=text name="Nomin" value="<%=Request("Nomin")%>" size=3 maxlength=2>��
		<!-- 2008/01/31 Add E G.Ariola -->
  </TD></TR>
  <TR>
    <TD><DIV class=bgb>��R���ԋp��</DIV></TD>
    <TD><INPUT type=text name="RPlace" value="<%=Request("RPlace")%>" size=30>
    </TD></TR>
<%'3th�ǉ� End%>
  <TR>
    <TD><DIV class=bgb>�ԋp�\������i�t���[�^�C���j</DIV></TD>
    <TD><INPUT type=text name="Rnissu" value="<%=Request("Rnissu")%>">
    </TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>���l�P</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�Q</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73></TD></TR>
<%'Del 3th  <TR>
'    <TD><DIV class=bgb>���l�R</DIV></TD>
'    <TD><INPUT type=text name="Comment3" value="<%=Request("Comment3")% >" size=13 maxlength=10></TD></TR>%>
<%'C-002 ADD End %>

   <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>�o�^�S����</DIV></TD>
  	<TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerSubName")%>" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>
  
<% If Mord=1 AND Request("UpFlag")<>1 Then %>
  <TR>
    <TD colspan=2 align=center>
    <DIV class=bgw>�w�����ւ̉񓚁@�@�@Yes�@�@�@�@�@</DIV>
    </TD></TR>
<% ElseIf Mord =2 Then %>
  <TR>
    <TD colspan=2 align=center>
    <DIV class=bgw>�w�����ւ̉񓚁@�@�@No�@�@�@�@�@</DIV>
    </TD></TR>
  <TR>
    <TD colspan=2 align=center>
       <DIV class=alert><B>�����Ӂ�</B>�񓚂�No�Ŏw��̏ꍇ�͓��͂����f�[�^�͔��f����܂���B</DIV>
    </TD></TR>
<% End If %>
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name=UpFlag value="<%=Request("UpFlag")%>" >
       <INPUT type=hidden name=UpUser  value="<%=Request("UpUser")%>">
       <INPUT type=hidden name="compFlag"  value="<%=Request("compFlag")%>">
       <INPUT type=hidden name=flag value="<%=Flag%>" >
       <INPUT type=hidden name=num value="<%=Num%>" >
       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
<% IF Num > 1 Then call Set_CONnum End If%>
<% If Not ret Then %>
       <P><DIV class=alert>
        �w�肳�ꂽ��ЃR�[�h�͑��݂��܂���B<BR>
       �u�߂�v�{�^�����������A�ē��͂��Ă��������B
       </DIV></P>
<% Else %>
       <INPUT type=submit value="�n�j" onClick="return GoEntry()">
       <INPUT type=hidden name=Mord value="<%=Mord%>" >
<% End If %>
       <INPUT type=submit value="�߂�" onClick="return GoBackT()">
    </TD></TR>
<!-- 2018/10/16 H.Yoshikawa Add-S -->
<% If Mord=0 AND ret Then %>
  <TR>
    <TD colspan=2 align=center>
<!-- 2019/01/29 H.Yoshikawa Upd-S -->
<!--   <div class=observ><INPUT type=checkbox name="ChkRule" value="1" ><a href="../download/download.asp?guide=���[�h�����ӎ���.pdf">���[�h�����ӎ���</a>�𗝉������S��Ƃ����炵�܂��B</div> -->
       <div class=observ>
			�i���񏑁j<BR>
			<INPUT type=checkbox name="ChkRule" value="1" ><a href="../download/download.asp?guide=���[�h�����ӎ���.pdf">�u���[�h�����ӎ����v</a>�����炵�A���S��Ƃ��s���܂��B<BR>
			�i�u���[�h�����ӎ����v����炸�ɋN�������Q�ɂ��܂��ẮA<BR>�^�[�~�i���͈�؂̐ӔC�𕉂��܂���̂ł������肢�܂��B�j
		</div>
<!-- 2019/01/29 H.Yoshikawa Upd-E -->
    </TD>
  </TR>
<% End If %>
<!-- 2018/10/16 H.Yoshikawa Add-E -->
</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
<SCRIPT language=JavaScript>
setParam(document.dmi030F);
</SCRIPT>
</BODY></HTML>