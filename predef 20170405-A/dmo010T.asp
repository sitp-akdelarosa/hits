<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo010T.asp				_/
'_/	Function	:�����o���ꗗ��ʃg�b�v		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-003 2003/08/22	��Ɣԍ��ł̌���_/
'_/     		:C-004 2008/01/29	KeyDown.js �� finit ��ǉ����܂���
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'���[�U�f�[�^����
        dim USER, COMPcd, LinUN
        USER   = Session.Contents("userid")
        COMPcd = Session.Contents("COMPcd")
	LinUN  = Session.Contents("LinUN")
	'2009/02/25 Add-S G.Ariola	
	Session("Key1") = ""
	Session("Key2") = ""
	Session("Key3") = ""
	
	Session("KeySort1") = ""
	Session("KeySort2") = ""
	Session("KeySort3") = ""
	'2009/02/25 Add-E G.Ariola	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�����o���ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<!--2008/01/29 Add-S M.Marquez-->
<SCRIPT src="./JS/KeyDown.js"></SCRIPT>
<!--2008/01/29 Add-E M.Marquez-->
<SCRIPT language=JavaScript>
<!--
//����
function Serch(){
  target = document.serch;
  if(!target.way[0].checked && !target.way[1].checked && !target.way[2].checked){
    alert("�����Ώۂ�I�����Ă�������");
    return false;
  }
  Num=LTrim(target.SortKye.value);
  if(Num.length==0){
    alert("��������ԍ����L�����Ă�������");
    target.SortKye.focus();
    return false;
  }
  if(!CheckEisu(target.SortKye.value)){
    alert("��������ԍ��ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕������L�����Ȃ��ł�������");
    target.SortKye.focus();
    return false;
  }
  if(target.way[0].checked){
    parent.DList.SerchC("3",target.SortKye.value);
  } else if(target.way[1].checked){
    parent.DList.SerchC("4",target.SortKye.value);
  } else {
    parent.DList.SerchC("11",target.SortKye.value);
  }

}
//�\�[�g
function sort(){

  target = document.serch;
  //alert(target.Sort.options[target.Sort.selectedIndex].value);
  target.SortFlag.value=target.Sort.options[target.Sort.selectedIndex].value;
  target.target="DList";
  target.action="./dmo010L.asp";
  target.submit();
}

//2008-01-29 Add-S M.Marquez
function finit(){
//    document.serch.Sort.focus();
document.serch.SortKye.focus();
}
//2008-01-29 Add-E M.Marquez
// -->

function OpenCodeWin()
{
	var CodeWin;
	var w=400;
	var h=300;
	var l=0;
	var t=0;
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
	
  CodeWin = window.open("./sort.asp?user=<%=Session.Contents("userid")%>&left_menu=1","codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);
  CodeWin.focus();

}

function showContent(){
    var target1 = document.getElementById("loading");
    target1.style.display='block';
    //show content	
    //parent.DList.document.getElementById("content").style.display='block';
}
</SCRIPT>
<style>
TD.bordering
{
    BORDER-BOTTOM: 1px dotted #000000;
    BORDER-LEFT: 1px dotted #000000;
    BORDER-RIGHT: 1px dotted #000000;
    BORDER-TOP: 1px dotted #000000;
	
}
</style>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="finit();">
<!-------------�����o���ꗗ���Top--------------------------->
<TABLE border=1 cellPadding=3 cellSpacing=0 align=center width="100%">
   <TR>
     <TD width="15%">���O�C�����[�U</TD>
     <TD><%=LinUN%></TD>
     <TD width="7%"><%=USER%></TD>
     <TD width="5%"><%=COMPcd%></TD></TR>
</TABLE>
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <FORM name="serch" action="">
   <TR><TD width="60%"><B class=title>�����o��ƈꗗ</B><BR><BR>
  
<%'CW-024 Dell IF SortFlag <> 6 Then %>

     <SELECT name="Sort" onChange="sort();">
       <OPTION value=0>���o�\������ɕ\��(�����ȍ~�̂�)</OPTION>
       <OPTION value=12>���o�\������ɕ\��(�����������\��)</OPTION>
<%'       <OPTION value=1>�ԋp��v����R���e�i���ɕ\��</OPTION>%>
       <OPTION value=2>�w���悪���񓚂̃R���e�i�ꗗ</OPTION>
       <OPTION value=7>�w����񓚂�No�̃R���e�i�ꗗ</OPTION>
       <OPTION value=6>���o�������������ׂĕ\��</OPTION>
       <OPTION value=5>�S���\��</OPTION>
     </SELECT>
<%'CW-024 Dell End If %>

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=button value="���בւ�" OnClick="OpenCodeWin()">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=button style="WIDTH: 120px;" value="�����o�f�[�^�X�V" OnClick="parent.parent.Top.location.reload(true);sort();">
<!--<div class="right" id="loading" >���΂炭���҂����������B&nbsp;<IMG border=0 src=Image/loaded.gif></div> -->
<TD align="left"  width="30%">
<!--2009/07/16 Upd-E G.Ariola -->
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
<TR><TD nowrap>
	<INPUT type=hidden name=SortFlag value="" >
	<INPUT type=radio name="way">�R���e�i�ԍ�
	<INPUT type=radio name="way">�a�k�ԍ�
	<INPUT type=radio name="way">��Ɣԍ�<BR>
	<INPUT type=text  name="SortKye" maxlength=20 size=27>
	<INPUT type=button value="����" onClick="Serch();">
</TD><TR>	
</TABLE>
	</TD>
<TD align="left"  width="10%">&nbsp;</TD>
	<TR>
  </FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
