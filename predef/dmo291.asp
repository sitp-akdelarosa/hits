<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo291.asp				_/
'_/	Function	:���O����o�w��������������		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:								_/
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
'���O�o��
  WriteLogH "b309", "����o�w�������", "01",""

'�T�[�o�����̎擾
  dim DayTime,day
  getDayTime DayTime
  day = DayTime(0) & "�N" & DayTime(1) & "��" & DayTime(2) & "��" 

'�O��ʂ���̃f�[�^�擾
  dim vanDate,vanTime,YY,i
  dim COMPcd1,vanMon,vanDay
  COMPcd1 = Request("COMPcd1")

'���̐��`
  vanMon =Right("00" & Request("vanMon"),2)
  vanDay =Right("00" & Request("vanDay"),2)
  If Request("vanHou")= "" Then
    vanTime=""
  Else
    vanTime=Right("00" & Request("vanHou"),2) & "��" & Right("00" & Request("vanMin"),2) & "��"
  End IF
  '���̔N�x������
  If DayTime(1) > vanMon Then	'���N
    YY = DayTime(0) +1
  ElseIf DayTime(1) = vanMon AND DayTime(2) > vanDay Then
    YY = DayTime(0) +1
  Else
    YY = DayTime(0)
  End If
  If vanMon = "00" Or vanDay = "00" Then
    vanDate= ""
  Else
    vanDate= YY &"�N"& vanMon &"��"& vanDay &"���@"& vanTime
  End If

  
'�Z�b�V�������烆�[�U���̂��擾
  Dim SjManN
  SjManN = Session.Contents("LinUN")

'DB����̃f�[�^�擾
  '�G���[�g���b�v�J�n
  on error resume next
  'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  '��ƎҖ��̎擾
  Dim WkManN
  If Trim(COMPcd1)="" OR COMPcd1=Null Then
    WkManN=SjManN
  Else
    StrSQL = "Select FullName From mUsers Where HeadCompanyCode='" & COMPcd1 &"'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "1","b309","01","����o�w������������E��ƎҖ��擾","102","SQL:<BR>"&strSQL
    end if
    WkManN= Trim(ObjRS("FullName"))
    ObjRS.close
  End If
'�w���ғd�b�ԍ��擾
  dim USER,TelNo
  USER       = Session.Contents("userid")
  StrSQL = "select TelNo from mUsers where UserCode='" & USER &"'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b309","01","����o�w������������E�w���ғd�b�ԍ��擾","102","SQL:<BR>"&strSQL
  end if
  TelNo = Trim(ObjRS("TelNo"))
  ObjRS.close
  If TelNo<>"" Then
    TelNo="�i�d�b�ԍ��F"&TelNo&"�j"
  End If

  'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�w�����������</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
//�w���������ʂ�
function GoNext(){
  target=document.dmo291F;
  newWin = window.open("", "Print2", "width=950,height=850,left=10,top=10,resizable=yes,scrollbars=yes,menubar=yes,top=0");
  target.target="Print2";
  target.submit();
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmo291F.WkManN.focus();
}
//2008-01-31 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onload="finit();">
<!-------------����o�w��������������--------------------------->
<FORM name="dmo291F" method="POST" action="./dmo292.asp";>
<CENTER><B class=titleB>����o�w����</B></CENTER>
<DIV style=text-align:right;>�쐬&nbsp;<%=day%></DIV>
<INPUT type=hidden name="day" value="<%=day%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR>
    <TD valign=top>�w����</TD><TD valign=top>��<%=SjManN%></TD>
    <TD>�i�S���ҁF�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�j<BR>
        <%=TelNo%></TD></TR>
  <TR>
    <TD>��Ǝ�</TD><TD>��<INPUT type=text name="WkManN" value="<%=WkManN%>"></TD><TD></TD></TR>
  <TR>
    <TD colspan=2>�u�b�L���O�ԍ��@�@�E�E�E�E�E�E</TD><TD><%=Request("BookNo")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="SjManN" value="<%=SjManN%>">
<INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>">
<INPUT type=hidden name="TelNo" value="<%=TelNo%>">

<TABLE border=0 cellPadding=0 cellSpacing=0 width=85% align=center>
  <TR><TD></TD><TD>�T�C�Y</TD><TD>�^�C�v</TD><TD>����</TD><TD>�ގ�</TD><TD>�s�b�N�ꏊ</TD><TD></TD><TD>�{��</TD></TR>
<% For i=0 To 4%>
  <TR><TD>(<%=i+1%>)</TD>
      <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4 maxlength=2>'</TD>
      <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4 maxlength=2></TD>
      <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4 maxlength=2></TD>
      <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4 maxlength=1></TD>
      <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 maxlength=20></TD>
      <TD>�E�E�E</TD>
      <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4   maxlength=3></TD></TR>
<% Next %>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>�P�D</TH>
    <TD><B>�u�b�L���O���&nbsp;</B></TD><TD></TD></TR>
  <TR>
    <TD>�i�D�Ёj</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>�i�D���j</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>�i�d���n�j</TD><TD><%=Request("delivTo")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="shipFact" value="<%=Request("shipFact")%>">
<INPUT type=hidden name="shipName" value="<%=Request("shipName")%>">
<INPUT type=hidden name="delivTo"  value="<%=Request("delivTo")%>">

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>�Q�D</TH>
    <TD><B>�o���l�ߏ��</B></TD><TD></TD></TR>
  <TR>
    <TD>�i�o���l�ߓ����j</TD><TD><%=vanDate%></TD></TR>
  <TR>
    <TD>�i�o���l�ߏꏊ�P�j&nbsp;</TD><TD><%=Request("vanPlace1")%></TD></TR>
  <TR>
    <TD>�i�o���l�ߏꏊ�Q�j</TD><TD><%=Request("vanPlace2")%></TD></TR>
  <TR>
    <TD>�i�i���j</TD><TD><%=Request("goodsName")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="vanDate"  value="<%=vanDate%>">
  <INPUT type=hidden name="vanPlace1" value="<%=Request("VanPlace1")%>">
  <INPUT type=hidden name="vanPlace2" value="<%=Request("VanPlace2")%>">
  <INPUT type=hidden name="goodsName" value="<%=Request("GoodsName")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>�R�D</TH>
    <TD><B>�������</B></TD><TD></TD></TR>
  <TR>
    <TD>�i������b�x�j</TD><TD><%=Request("Terminal")%></TD></TR>
  <TR>
    <TD>�i�b�x�J�b�g���j&nbsp;</TD><TD><%=Request("CYCut")%></TD></TR>
</TABLE><P>
  <INPUT type=hidden name="Terminal"  value="<%=Request("Terminal")%>">
  <INPUT type=hidden name="CYCut"    value="<%=Request("CYCut")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>�S�D</TH>
    <TD><B>���l</B></TD><TD></TD></TR>
  <TR>
    <TD>�i���l�P�j&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD>�i���l�Q�j</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="Comment1"  value="<%=Request("Comment1")%>">
<INPUT type=hidden name="Comment2"  value="<%=Request("Comment2")%>">
<CENTER>
  <INPUT type=button value="�n�j" onClick="GoNext()">
  <INPUT type=button value="����" onClick="window.close()">
</CENTER>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
