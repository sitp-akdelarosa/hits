<%@LANGUAGE = VBScript%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo010B.asp				_/
'_/	Function	:�����o���ꗗ��ʃt�b�^		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-001 2003/07/29	CSV�o�͑Ή�	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�����o���ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//�Ɖ��
function GoSyokaizumi(){
  try{
    parent.DList.GoSyokaizumi();
  }catch(e){}
}
//CSV
function GoCSV(){
  try{
    parent.DList.GoCSV();
  }catch(e){}
}

<%'//�����ʕ\��
'function GoPlint(){
'  Win = window.open('', 'PrintWin', 'width=1000,height=600,menubar=yes,scrollbars=yes');
'  document.next.target="PrintWin";
'  document.next.action="./dmo090.asp";
'  document.next.submit();
'}
%>
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�����o���ꗗ���Bottom--------------------------->
<CENTER>
  <FORM name="next" action="">
    <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" height=35>
    <TR><TD>
        <A HREF='JavaScript:GoSyokaizumi()'>�w�����</A>�E�E�E�\������Ă���S�Ă̖��񓚃f�[�^�̉񓚂��uYes�v�ɂ��܂��B
        </TD>
        <TD>
        <A HREF="JavaScript:GoHelp(1)">�w���v</A>�E�E�E��ʓ��̋@�\�̐�����ʂ�\�����܂��B
        </TD></TR>
    <TR><TD colspan=2>
<!--        <A HREF="JavaScript:GoPlint()">�����ʕ\��</A>�E�E�E�\�����e������ɓK������ʂŕ\�����܂��B�a�k�ԍ��w���ꗗ����I�����ꂽ���̂́A���Y����R���e�i�ɓW�J���ĕ\�����܂��B -->
        <A HREF="JavaScript:GoCSV()">CSV�t�@�C���o��</A>�E�E�E�\�����e��CSV�t�@�C���ɏo�͂��܂��B�a�k�ԍ��w���ꗗ����I�����ꂽ���̂́A���Y����R���e�i�ɓW�J���ĕ\�����܂��B
        <INPUT type=hidden name="SortFlag" value="">
        </TD></TR>
    </TABLE>
  </FORM>
</CENTER>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
