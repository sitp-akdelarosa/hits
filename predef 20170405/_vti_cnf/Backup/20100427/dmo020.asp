<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo020.asp				_/
'_/	Function	:���O�����o���͉��(�\��)		_/
'_/	Date		:2003/05/27				_/
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

'�f�[�^�擾
  dim SakuNo,Flag,UpFlag,Num,CONnumA(),CMPcd(5),Rmon,Rday
  dim param,i,j,Way
  Way   =Array("","�w�肠��","�w��Ȃ�","�ꗗ����I��","�a�k�ԍ�")
  SakuNo= Request("SakuNo")
  Flag= Request("flag")
  WriteLogH "b10"&(2+Flag), "�����o���O������("&Way(Flag)&")", "11",""
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

'�\����������
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
<TITLE>�����o������(�\��)</TITLE>
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

//�R���e�i�ڍ׉��
function GoConInfo(){
  target=document.dmo020F;
  ConInfo(target,<%=Flag%>,0);
  return false;
}
//�X�V��ʂ�
function GoReEntry(){
  target=document.dmo020F;
  target.action="./dmi021.asp";
  return true;
}
//�w�������������ʂ�
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
<!-------------�����o������(�\��)���--------------------------->
<FORM name="dmo020F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD><B>�����o������(�\�����[�h)</B></TD>
    <TD><TABLE border=1 cellPadding=3 cellSpacing=0 align="right">
          <TR bgcolor="#f0f0f0"><TD>��Ɣԍ�</TD><TD><%=SakuNo%></TD></TR>
        </TABLE>
    </TD></TR>
  <TR>
<% If Flag=4 Then %>
    <TD><DIV class=bgb>�a�k�m���D</DIV></TD>
    <TD><INPUT type=text name="BLnum" value="<%=Request("BLnum")%>">�@�@<%=Way(Flag)%>
    <INPUT type=hidden name="CONnum" value="<%=CONnumA(0)%>"></TD></TR>
<% Else %>
    <TD><DIV class=bgb>�R���e�i�m���D</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnumA(0)%>">�@�@<%=Way(Flag)%></TD></TR>
    <INPUT type=hidden name="BLnum" value="<%=Request("BLnum")%>"></TD></TR>
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
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD Align=right>�w�����S����</TD>
    <TD>
        <INPUT type=text name="SubName" readonly = "readonly" value="<%=Request("TruckerSubName")%>" maxlength=16>
    </TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
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
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" size=40></TD></TR>
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
<!-- 2009/03/10 R.Shibuta Add-S -->
  <TR>
   <TD><DIV class=bgy>�o�^�S����</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>

<%'Del 3th  <TR>
'    <TD><DIV class=bgb>���l�R</DIV></TD>
'    <TD><INPUT type=text name="Comment3" value="<%=Request("Comment3")% >" size=13 maxlength=10></TD></TR>%>
<%'C-002 ADD End %>
  <TR>
    <TD colspan=2 align=center>
       <DIV class=alert><B>�����Ӂ�</B>�D�Ђɂ���Ă̓Q�[�g�ŔF��ID�̓��͂��K�v�ɂȂ�܂�</DIV>
    </TD></TR>
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name="UpUser" value="<%=Request("UpUser")%>" >
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
 <!-- 2009/08/04 Tanaka Add-S -->
       <INPUT type=hidden name="TruckerName" value="<%=Request("TruckerName")%>" >
 <!-- 2009/08/04 Tanaka Add-E -->
<% If Num > 1 Then call Set_CONnum End If%>
       <INPUT type=button value="�w�������" onClick="GoSijiPrint()">
       <INPUT type=hidden name="SakuNo" value="<%=SakuNo%>">
       <INPUT type=hidden name=flag value="<%=Flag%>" >
       <INPUT type=hidden name=num value="<%=Num%>" >
<%'20030909 IF Request("compFlag") AND (UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1) Then %>

       <!--INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>"-->
<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
' IF UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1 Then 
  IF UCase(Session.Contents("userid"))=CMPcd(0) Or (Request("compFlag") AND Request("TruckerFlag")<>1) Then %>
       <INPUT type=submit value="�X�V���[�h" onClick="return GoReEntry()">
       <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
       <INPUT type=hidden name="WkCNo"    value="<%=Request("WkCNo")%>">
       <INPUT type=hidden name="TruckerFlag" value="<%=Request("TruckerFlag")%>">
<%End IF%>
       <INPUT type=submit value="����" onClick="window.close()">
       <P>
       <INPUT type=submit value="�R���e�i���" onClick="return GoConInfo()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
<SCRIPT language=JavaScript>
setParam(document.dmo020F);
</SCRIPT>
</BODY></HTML>
