<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits											_/
'_/	FileName	:dmo312.asp										_/
'_/	Function	:���O���������擾�ABooking�������J�E���g��	_/
'_/				:		�����Ȃ��dmi314.asp					_/
'_/				:		�P���Ȃ��ShoriMode�Ő���				_/
'_/				:		"FLin"������		dmi315.asp			_/
'_/				:		"EMoutInf"����o���BookInfo()			_/
'_/				:		"EMoutUpd"����o�o�^dmi215.asp			_/
'_/	Date		:2006/03/06										_/
'_/	Code By		:SEIKO Electric.Co ���c�E�l						_/
'_/	Modify		:												_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>

<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  
'�f�[�^���擾
  dim CONnum,BookNo,User,ShipLine,ShoriMode,Mord
  dim WkNo													'2016/10/18 H.Yoshikawa Add
  BookNo = Trim(Request("BookNo"))
  CONnum = Trim(Request("CONnum"))
  ShoriMode = Trim(Request("ShoriMode"))
  ShipLine = Trim(Request("ShipLine"))
  Mord = Trim(Request("Mord"))
  User   = Session.Contents("userid")
  WkNo = gfTrim(Request("WkNo"))							'2016/10/18 H.Yoshikawa Add
'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL, ErrerM
  ConnDBH ObjConn, ObjRS
  
'2016/10/18 H.Yoshikawa Add Start
'��Ɣԍ��w��̏ꍇ�́ACYVanInfo���BookNo��ShipLine���擾
  ret="0"
  if WkNo <> "" then
	StrSQL = "select BookNo, ShipLine "&_
			 "from CYVanInfo "&_
			 "where WkNo = '"& gfSQLEncode(WkNo) & "' "
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB�ؒf
		jampErrerP "1","b401","01","���O�o�^�F��Ɣԍ����݃`�F�b�N","101","SQL:<BR>"&StrSQL
	end if
	if ObjRS.eof then
		ret="3"
		ErrerM="�w�肵����Ɣԍ����V�X�e���ɓo�^����Ă��܂���B<BR>���͂̊ԈႢ���Ȃ����ԍ����m�F���Ă��������B</P>"
	else
		BookNo = gfTrim(ObjRS("BookNo"))
		ShipLine = gfTrim(ObjRS("ShipLine"))
	end if
	ObjRS.Close
  end if
if ret = "0" then
'2016/10/18 H.Yoshikawa Add End

'�u�b�L���O�ԍ��̑��݃`�F�b�N,ret=1:Booking0��(���̓G���[),ret=0:Booking�P��,ret=2:BookingN��
  dim dummy,ret
  ret="0"
  StrSQL = "select count(BOK.BookNo) as Num "&_
		   ",max( BOK.ShipLine) as ShipLine "&_
		   "from(select distinct BookNo,shipline from Booking) as BOK "&_
		   "where BOK.BookNo='"& BookNo & "' "
	if ShipLine<>"" then
		strsql=strsql & "and BOK.ShipLine='"& ShipLine & "' "
	end if
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b401","01","���O�o�^�F�d���`�F�b�N","101","SQL:<BR>"&StrSQL
  end if
  If Trim(ObjRS("Num")) = "0" Then
    ret="1"
    ErrerM="�w�肵���u�b�L���ONo���V�X�e���ɓo�^����Ă��܂���B<BR>���͂̊ԈႢ���Ȃ����ԍ����m�F���Ă��������B</P>"
  ElseIf Trim(ObjRS("Num")) > "1" then
    ret="2"
    ErrerM="���͂��ꂽBookin�ԍ��͕����o�^����Ă��܂��B</P>"
  else
    ShipLine = Trim(ObjRS("ShipLine"))
  End If
  ObjRS.Close
end if									'2016/10/18 H.Yoshikawa Add
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0
  
  If ret ="0"Then
    WriteLogH "b402", "���������O������", "01",BookNo&",���͓��e�̐���:0(������)"
  elseif ret="2" then
    WriteLogH "b402", "���������O������", "01",BookNo&",���͓��e�̐���:2(�������A������)"
  Else
    WriteLogH "b402", "���������O������", "01",BookNo&",���͓��e�̐���:1(���)"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�D��БI��</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){

<% If ret="0" Then %>
	<% If ShoriMode="FLin" Then %>
		window.resizeTo(850,690);
		target.action="./dmi315.asp";
		target.submit();
	<% elseIf ShoriMode="EMoutUpd" Then %>
		target.action="./dmi215.asp";
		target.submit();
	<% elseIf ShoriMode="EMoutInf" Then %>
		BookInfo(target);
	<% End If %>
<% elseIf ret="2" Then %>
  window.resizeTo(500,500);
  window.focus();
<% elseIf ret="1" Then %>
  window.resizeTo(500,500);
  window.focus();
<% End If %>
}

function GoNext(){
  target=document.dmi312F;
	document.dmi312F.action="./dmi314.asp";
    Num=LTrim(target.ShipLine.value );
    if(Num.length==0){
      alert("�D�Ђ̓��������L�����Ă�������");
      target.ShipLine.focus();
      return;
    }

  if(!CheckEisu(target.ShipLine.value)){
    alert("�������ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕������L�����Ȃ��ł�������");
    target.ShipLine.focus();
    return;
  }
  chengeUpper(target);
  target.submit();
}
function GoBack(){
	<% If ShoriMode="FLin" Then %>
		target=document.dmi312F;
		target.action="./dmi310.asp";
	<% ElseIf ShoriMode="EMoutUpd" Then %>
      	window.open('dmi210.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
	<% Else %>
		target=document.dmi312F;
		window.close(); 
	<% End If %>
	target.submit();
}

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="setParam(document.dmi312F)">
<FORM name="dmi312F" method="POST">
<!-------------���������擾���--------------------------->
<% If ret="0" Then %>

	<INPUT type=hidden name="BookNo" value="<%=BookNo%>">
	<INPUT type=hidden name="CONnum" value="<%=CONnum%>">
	<INPUT type=hidden name="ShipLine" value="<%=ShipLine%>">
	<INPUT type=hidden name="Mord" value="<%=Mord%>">
	<INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
	<INPUT type=hidden name="WkNo" value="<%=WkNo%>">						<!-- 2016/10/18 H.Yoshikawa Add -->


<% ElseIf ret="2" Then %>
<CENTER>
	<br><br><br>
  <DIV class=alert>
    <%= ErrerM%>
  </DIV>
	<table>
		<TD colspan="3" align="center">
			<br><br>
			<B>�D�Ђ̓������𔼊p�A���t�@�x�b�g1�������͂��n�j�������Ă��������B</B><BR>
			<br><br>
			<INPUT type=text  name="ShipLine" maxlength=1 size=3><BR>
			<INPUT type=hidden name="BookNo" value="<%=BookNo%>">
			<INPUT type=hidden name="CONnum" value="<%=CONnum%>">
			<INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
			<INPUT type=hidden name="Mord" value="<%=Mord%>">
			<br><br>
			<P><INPUT id=button1 type=button value="�@�߂�@" 
				name=button1 LANGUAGE=javascript onclick="GoBack()">
			<INPUT id=button1 type=button value="�@�n�j�@" 
				name=button1 LANGUAGE=javascript onclick="GoNext()"></P>
		</TD>
		</table>
</CENTER>

<% Else %>
<CENTER>
  <DIV class=alert>
    <%= ErrerM %>
  </DIV>
  <P><INPUT type=button value="����" onClick="window.close()" id=button1 name=button1></P>
</CENTER>

<% End If %>

  <INPUT type=hidden name=DataNum value="<%=Request("Num")%>">
  <INPUT type=hidden name=SortFlag value="<%=Request("SortFlag")%>" >
  <INPUT type=hidden name=SortKye value="<%=Request("SortKye")%>" >
  <INPUT type=hidden name=CompF value="<%=Request("CompF")%>" >
  <INPUT type=hidden name=COMPcd0 value="<%=Request("COMPcd0")%>" >
  <INPUT type=hidden name=COMPcd1 value="<%=Request("COMPcd1")%>" >
  <INPUT type=hidden name=strWhere value="<%=Request("strWhere")%>">

  

</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>

